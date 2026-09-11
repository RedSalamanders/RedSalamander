#if defined(ENABLE_TESTS)

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <winsock2.h>
#include <ws2tcpip.h>

#include "FileSystemCurl.Internal.h"
#include "FileSystemRouteContract.h"
#include "FileOperationTraversalPolicy.h"
#include "PlugInterfaces/Informations.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <charconv>
#include <chrono>
#include <condition_variable>
#include <cstdio>
#include <cstdint>
#include <format>
#include <iterator>
#include <limits>
#include <map>
#include <mutex>
#include <optional>
#include <set>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <utility>
#include <vector>

namespace FileSystemCurlInternal
{
namespace
{
using namespace std::chrono_literals;

constexpr Common::DebugSelfTest::Check DebugCheck{L"FileSystemCurl FTP transport"};
constexpr std::string_view kUploadSiblingMarker   = ".redsalamander-upload-";
constexpr std::string_view kRollbackSiblingMarker = ".redsalamander-rollback-";
constexpr std::string_view kWriterSiblingMarker   = ".redsalamander-writer-";
constexpr size_t kMaximumDataTransferBytes        = 1024u * 1024u;
constexpr auto kSocketPollInterval                = 50ms;
constexpr auto kControlIdleTimeout                = 30s;
constexpr auto kDataOperationTimeout              = 5s;

enum class UploadRetention
{
    Complete,
    StrictPrefix,
    FailAfterStore,
};

enum class DownloadDelivery
{
    Complete,
    StrictPrefix,
    Overlong,
};

enum class LateListingReply
{
    Complete,
    Refuse,
    WaitForCancel,
};

struct CommandRecord final
{
    std::string verb;
    std::string argument;
};

enum class DataTransferKind : size_t
{
    List,
    NamesOnly,
    Retrieve,
    Store,
    Count,
};

struct DataTransferCounts final
{
    uint64_t firstUse = 0u;
    uint64_t reuse    = 0u;
};

struct EndpointSnapshot final
{
    std::map<std::string, std::vector<uint8_t>> files;
    std::set<std::string> directories;
    std::vector<CommandRecord> commands;
    std::vector<std::string> deletedPaths;
    size_t lastUploadReceivedBytes            = 0u;
    size_t lastUploadRetainedBytes            = 0u;
    size_t lastDownloadSentBytes              = 0u;
    size_t beforeRenameInjectionCount         = 0u;
    size_t afterRenameInjectionCount          = 0u;
    size_t beforeListInjectionCount           = 0u;
    size_t afterDirectoryCreateInjectionCount = 0u;
    size_t rejectedDeleteCount                = 0u;
    size_t lostMutationReplyCount             = 0u;
    size_t rejectedDirectoryAccessCount       = 0u;
    size_t lateListingReachedCount            = 0u;
    size_t lateListingReleasedCount           = 0u;
    size_t lateListingTimeoutCount            = 0u;
    size_t passiveListenerCreations           = 0u;
    size_t passiveListenerReuses              = 0u;
    size_t passivePendingRetirements          = 0u;
    size_t idleControlClosures                = 0u;
    std::array<DataTransferCounts, static_cast<size_t>(DataTransferKind::Count)> dataTransfers{};
    uint64_t listPayloadBuildUs = 0u;
    uint64_t listPayloadBytes   = 0u;
    HRESULT serverHr            = S_OK;
};

[[nodiscard]] HRESULT SocketErrorToHResult() noexcept
{
    const int error = WSAGetLastError();
    return error == 0 ? E_FAIL : HRESULT_FROM_WIN32(static_cast<unsigned long>(error));
}

[[nodiscard]] HRESULT ConfigureSocketTimeouts(SOCKET socketValue) noexcept
{
    constexpr DWORD kTimeoutMs = 2000u;
    if (setsockopt(socketValue, SOL_SOCKET, SO_RCVTIMEO, reinterpret_cast<const char*>(&kTimeoutMs), static_cast<int>(sizeof(kTimeoutMs))) != 0)
    {
        return SocketErrorToHResult();
    }
    if (setsockopt(socketValue, SOL_SOCKET, SO_SNDTIMEO, reinterpret_cast<const char*>(&kTimeoutMs), static_cast<int>(sizeof(kTimeoutMs))) != 0)
    {
        return SocketErrorToHResult();
    }
    return S_OK;
}

enum class LoopbackEndpointMode
{
    Listening,
    BoundOnly,
};

[[nodiscard]] HRESULT CreateLoopbackEndpoint(wil::unique_socket& listenerOut,
                                             unsigned short& portOut,
                                             LoopbackEndpointMode mode = LoopbackEndpointMode::Listening) noexcept
{
    listenerOut.reset();
    portOut = 0u;

    wil::unique_socket listener(socket(AF_INET, SOCK_STREAM, IPPROTO_TCP));
    if (! listener)
    {
        return SocketErrorToHResult();
    }

    // Control/data listeners and the refusal oracle must own their loopback
    // port exclusively. Never allow a same-user reuse bind to share a fixture.
    const BOOL exclusive = TRUE;
    if (setsockopt(listener.get(), SOL_SOCKET, SO_EXCLUSIVEADDRUSE, reinterpret_cast<const char*>(&exclusive), static_cast<int>(sizeof(exclusive))) != 0)
    {
        return SocketErrorToHResult();
    }

    sockaddr_in address{};
    address.sin_family      = AF_INET;
    address.sin_addr.s_addr = htonl(INADDR_LOOPBACK);
    address.sin_port        = 0u;

    if (bind(listener.get(), reinterpret_cast<const sockaddr*>(&address), static_cast<int>(sizeof(address))) != 0)
    {
        return SocketErrorToHResult();
    }
    // BoundOnly retains ownership of the port while deterministically refusing
    // connections; a released-port test could accidentally reach another server.
    if (mode == LoopbackEndpointMode::Listening && listen(listener.get(), SOMAXCONN) != 0)
    {
        return SocketErrorToHResult();
    }

    int addressBytes = static_cast<int>(sizeof(address));
    if (getsockname(listener.get(), reinterpret_cast<sockaddr*>(&address), &addressBytes) != 0)
    {
        return SocketErrorToHResult();
    }

    portOut = ntohs(address.sin_port);
    if (portOut == 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    listenerOut = std::move(listener);
    return S_OK;
}

[[nodiscard]] HRESULT WaitForReadable(SOCKET socketValue, std::stop_token stopToken, std::chrono::steady_clock::time_point deadline) noexcept
{
    while (! stopToken.stop_requested())
    {
        const auto now = std::chrono::steady_clock::now();
        if (now >= deadline)
        {
            return HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT);
        }

        const auto remaining = deadline - now;
        const auto wait      = (std::min)(std::chrono::duration_cast<std::chrono::microseconds>(remaining),
                                          std::chrono::duration_cast<std::chrono::microseconds>(kSocketPollInterval));
        timeval timeout{};
        timeout.tv_sec  = static_cast<long>(wait.count() / 1000000ll);
        timeout.tv_usec = static_cast<long>(wait.count() % 1000000ll);

        fd_set readSet;
        FD_ZERO(&readSet);
        FD_SET(socketValue, &readSet);
        const int selectResult = select(0, &readSet, nullptr, nullptr, &timeout);
        if (selectResult > 0)
        {
            return S_OK;
        }
        if (selectResult < 0)
        {
            return SocketErrorToHResult();
        }
    }

    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

[[nodiscard]] HRESULT SendAll(SOCKET socketValue, std::string_view bytes) noexcept
{
    size_t offset = 0u;
    while (offset < bytes.size())
    {
        const size_t remaining = bytes.size() - offset;
        const int request      = static_cast<int>((std::min)(remaining, static_cast<size_t>((std::numeric_limits<int>::max)())));
        const int sent         = send(socketValue, bytes.data() + offset, request, 0);
        if (sent == SOCKET_ERROR)
        {
            return SocketErrorToHResult();
        }
        if (sent == 0)
        {
            return HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED);
        }
        offset += static_cast<size_t>(sent);
    }
    return S_OK;
}

[[nodiscard]] HRESULT ReadControlLine(SOCKET socketValue, std::stop_token stopToken, std::string& pending, std::string& lineOut) noexcept
{
    lineOut.clear();
    const auto deadline = std::chrono::steady_clock::now() + kControlIdleTimeout;

    for (;;)
    {
        const size_t end = pending.find("\r\n");
        if (end != std::string::npos)
        {
            lineOut.assign(pending, 0u, end);
            pending.erase(0u, end + 2u);
            return S_OK;
        }

        if (pending.size() > 16u * 1024u)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }

        const HRESULT waitHr = WaitForReadable(socketValue, stopToken, deadline);
        if (FAILED(waitHr))
        {
            return waitHr;
        }

        std::array<char, 2048u> buffer{};
        const int received = recv(socketValue, buffer.data(), static_cast<int>(buffer.size()), 0);
        if (received == 0)
        {
            return HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED);
        }
        if (received == SOCKET_ERROR)
        {
            const int error = WSAGetLastError();
            if (error == WSAETIMEDOUT || error == WSAEWOULDBLOCK)
            {
                continue;
            }
            return HRESULT_FROM_WIN32(static_cast<unsigned long>(error));
        }
        pending.append(buffer.data(), static_cast<size_t>(received));
    }
}

[[nodiscard]] std::string UpperAscii(std::string_view value)
{
    std::string result(value);
    for (char& ch : result)
    {
        ch = static_cast<char>(std::toupper(static_cast<unsigned char>(ch)));
    }
    return result;
}

[[nodiscard]] std::string NormalizeFtpPath(std::string_view value, std::string_view currentDirectory)
{
    std::string path(value);
    while (! path.empty() && (path.front() == ' ' || path.front() == '\t'))
    {
        path.erase(path.begin());
    }
    while (! path.empty() && (path.back() == ' ' || path.back() == '\t'))
    {
        path.pop_back();
    }

    if (path.empty())
    {
        return currentDirectory.empty() ? std::string("/") : std::string(currentDirectory);
    }

    if (path.front() != '/')
    {
        std::string absolute(currentDirectory.empty() ? "/" : currentDirectory);
        if (absolute.back() != '/')
        {
            absolute.push_back('/');
        }
        absolute.append(path);
        path = std::move(absolute);
    }

    while (path.size() > 1u && path.back() == '/')
    {
        path.pop_back();
    }
    return path;
}

// Borrowed lexical slices of normalized UTF-8 fixture paths, not production
// UTF-16 route identity helpers. Callers retain the input and any state lock.
[[nodiscard]] constexpr std::string_view FtpFixtureLeafView(std::string_view path) noexcept
{
    const size_t slash = path.find_last_of('/');
    return slash == std::string_view::npos ? path : path.substr(slash + 1u);
}

[[nodiscard]] constexpr std::string_view FtpFixtureParentView(std::string_view path) noexcept
{
    if (path.empty() || path == "/")
    {
        return "/";
    }

    const size_t slash = path.find_last_of('/');
    return slash == std::string_view::npos || slash == 0u ? std::string_view("/") : path.substr(0u, slash);
}

static_assert(FtpFixtureParentView("") == "/");
static_assert(FtpFixtureParentView("/") == "/");
static_assert(FtpFixtureParentView("leaf.bin") == "/");
static_assert(FtpFixtureParentView("/chosen/leaf.bin") == "/chosen");
static_assert(FtpFixtureLeafView("").empty());
static_assert(FtpFixtureLeafView("/").empty());
static_assert(FtpFixtureLeafView("leaf.bin") == "leaf.bin");
static_assert(FtpFixtureLeafView("/chosen/literal%20 -> \xC3\xA9.bin") == "literal%20 -> \xC3\xA9.bin");
static_assert([]()
{
    std::string path = "/chosen/sub";
    path             = FtpFixtureParentView(path); // The CDUP caller assigns its own prefix.
    return path == "/chosen";
}());

class FakeFtpEndpoint final
{
public:
    explicit FakeFtpEndpoint(UploadRetention retention,
                             bool sizeCommandSupported         = true,
                             bool listIncludesSize             = true,
                             bool failWriterPromotion          = false,
                             DownloadDelivery downloadDelivery = DownloadDelivery::Complete) noexcept
        : _retention(retention),
          _sizeCommandSupported(sizeCommandSupported),
          _listIncludesSize(listIncludesSize),
          _failWriterPromotion(failWriterPromotion),
          _downloadDelivery(downloadDelivery)
    {
    }

    ~FakeFtpEndpoint() noexcept
    {
        Stop();
    }

    FakeFtpEndpoint(const FakeFtpEndpoint&)            = delete;
    FakeFtpEndpoint& operator=(const FakeFtpEndpoint&) = delete;
    FakeFtpEndpoint(FakeFtpEndpoint&&)                 = delete;
    FakeFtpEndpoint& operator=(FakeFtpEndpoint&&)      = delete;

    void SeedFile(std::string path, std::vector<uint8_t> bytes)
    {
        const std::string normalizedPath = NormalizeFtpPath(path, "/");
        std::scoped_lock lock(_stateMutex);
        AddDirectoryWithParentsLocked(FtpFixtureParentView(normalizedPath));
        _files[normalizedPath] = std::move(bytes);
    }

    void SeedDirectory(std::string path)
    {
        const std::string normalizedPath = NormalizeFtpPath(path, "/");
        std::scoped_lock lock(_stateMutex);
        AddDirectoryWithParentsLocked(normalizedPath);
    }

    // After the next successful RETR, insert this file so a later source LIST
    // (Move's delete pass) can see a name that preflight never committed.
    void SeedFileAfterNextRetrieve(std::string path, std::vector<uint8_t> bytes)
    {
        const std::string normalizedPath = NormalizeFtpPath(path, "/");
        std::scoped_lock lock(_stateMutex);
        _fileAfterNextRetrievePath    = normalizedPath;
        _fileAfterNextRetrieveBytes   = std::move(bytes);
        _fileAfterNextRetrievePending = true;
    }

    void SetDownloadDelivery(DownloadDelivery delivery) noexcept
    {
        std::scoped_lock lock(_stateMutex);
        _downloadDelivery = delivery;
    }

    void SetListedSize(std::string path, size_t listedSize)
    {
        const std::string normalizedPath = NormalizeFtpPath(path, "/");
        std::scoped_lock lock(_stateMutex);
        _listedSizes.insert_or_assign(normalizedPath, listedSize);
    }

    void SetSizeAuthenticationFailure(bool fail) noexcept
    {
        _sizeAuthenticationFailure.store(fail, std::memory_order_release);
    }

    void SetSizeMissingFailure(bool fail) noexcept
    {
        _sizeMissingFailure.store(fail, std::memory_order_release);
    }

    void SetFileBeforeRename(std::string path, std::vector<uint8_t> bytes)
    {
        const std::string normalizedPath = NormalizeFtpPath(path, "/");
        std::scoped_lock lock(_stateMutex);
        _filesBeforeRename.insert_or_assign(normalizedPath, std::move(bytes));
    }

    void SetFileAfterRename(std::string path, std::vector<uint8_t> bytes)
    {
        const std::string normalizedPath = NormalizeFtpPath(path, "/");
        std::scoped_lock lock(_stateMutex);
        _filesAfterRename.insert_or_assign(normalizedPath, std::move(bytes));
    }

    void SetLostMutationReply(std::string verb, std::string sourcePath, std::vector<uint8_t> replacement)
    {
        std::scoped_lock lock(_stateMutex);
        _lostMutationReplyVerb = std::move(verb);
        _lostMutationReplyPath = NormalizeFtpPath(sourcePath, "/");
        _lostMutationReplyReplacement = std::move(replacement);
    }

    // R0f-Curl witness: the fixture holds the reply to `verb` for `stallMs` (or until it stops),
    // which is what a server that stopped answering looks like to the provider.
    void SetStallVerb(std::string verb, unsigned int stallMs)
    {
        std::scoped_lock lock(_stateMutex);
        _stallVerb = std::move(verb);
        _stallMs   = stallMs;
    }

    void SetFileBeforeDirectoryList(std::string directory, std::string path, std::vector<uint8_t> bytes)
    {
        const std::string normalizedDirectory = NormalizeFtpPath(directory, "/");
        const std::string normalizedPath      = NormalizeFtpPath(path, "/");
        std::scoped_lock lock(_stateMutex);
        _filesBeforeDirectoryList.insert_or_assign(
            normalizedDirectory, PendingFileInjection{.path = normalizedPath, .bytes = std::move(bytes)});
    }

    void SetFileAfterDirectoryCreate(std::string directory, std::string path, std::vector<uint8_t> bytes)
    {
        const std::string normalizedDirectory = NormalizeFtpPath(directory, "/");
        const std::string normalizedPath      = NormalizeFtpPath(path, "/");
        std::scoped_lock lock(_stateMutex);
        _filesAfterDirectoryCreate.insert_or_assign(
            normalizedDirectory, PendingFileInjection{.path = normalizedPath, .bytes = std::move(bytes)});
    }

    void SetDeleteFailure(std::string path)
    {
        const std::string normalizedPath = NormalizeFtpPath(path, "/");
        std::scoped_lock lock(_stateMutex);
        _deleteFailures.insert(normalizedPath);
    }

    void SetDirectoryAccessDenied(std::string path)
    {
        std::scoped_lock lock(_stateMutex);
        _deniedDirectories.insert(NormalizeFtpPath(path, "/"));
    }

    void SetDirectoryRenameRefused(std::string path)
    {
        std::scoped_lock lock(_stateMutex);
        _refusedDirectoryRename = NormalizeFtpPath(path, "/");
    }

    void SetDirectoryPayload(std::string directory, std::string payload)
    {
        std::scoped_lock lock(_stateMutex);
        _directoryPayloads.insert_or_assign(NormalizeFtpPath(directory, "/"), std::move(payload));
    }

    void SetFailedListFinalReply(bool fail)
    {
        std::scoped_lock lock(_stateMutex);
        _failListFinalReply = fail;
    }

    // Transport-local ordering witness, not UI polling: wait for an exact prior DELE
    // before answering a later LIST. The wait never owns the command-serialization lock.
    void SetLateListingGate(std::string directory, std::string deletedPath, LateListingReply reply)
    {
        std::scoped_lock lock(_stateMutex);
        _lateListingDirectory   = NormalizeFtpPath(directory, "/");
        _lateListingDeletedPath = NormalizeFtpPath(deletedPath, "/");
        _lateListingReply       = reply;
    }

    [[nodiscard]] bool WaitForLateListingRelease(std::chrono::milliseconds timeout)
    {
        std::unique_lock lock(_stateMutex);
        return _lateListingChanged.wait_for(lock, timeout, [this]() noexcept { return _lateListingReleasedCount != 0u || _lateListingTimeoutCount != 0u; }) &&
               _lateListingReleasedCount != 0u;
    }

    [[nodiscard]] bool WaitForLateListingReached(std::chrono::milliseconds timeout)
    {
        std::unique_lock lock(_stateMutex);
        return _lateListingChanged.wait_for(lock, timeout, [this]() noexcept { return _lateListingReachedCount != 0u; });
    }

    void SetDeleteFailureForMarker(std::string marker)
    {
        std::scoped_lock lock(_stateMutex);
        _deleteFailureMarkers.insert(std::move(marker));
    }

    [[nodiscard]] HRESULT Start()
    {
        HRESULT hr = CreateLoopbackEndpoint(_listener, _port);
        if (FAILED(hr))
        {
            return hr;
        }

        _thread = std::jthread([this](std::stop_token stopToken) noexcept
        {
            try
            {
                ServerMain(stopToken);
            }
            catch (const std::bad_alloc&)
            {
                std::terminate();
            }
            catch (const std::exception&)
            {
                // Mandatory thread boundary: convert an unexpected fixture exception into deterministic test failure.
                Debug::Error(L"FileSystemCurl fake FTP endpoint terminated after std::exception.");
                SetServerFailure(E_FAIL);
            }
        });
        return S_OK;
    }

    void Stop() noexcept
    {
        if (_thread.joinable())
        {
            _thread.request_stop();
            _thread = std::jthread{};
        }
        std::vector<std::jthread> clients;
        {
            std::scoped_lock lock(_clientsMutex);
            clients.swap(_clients);
        }
        for (std::jthread& client : clients)
        {
            client.request_stop();
        }
        clients.clear();
        _listener.reset();
    }

    [[nodiscard]] unsigned short Port() const noexcept
    {
        return _port;
    }

    [[nodiscard]] EndpointSnapshot Snapshot() const
    {
        std::scoped_lock lock(_stateMutex);
        return EndpointSnapshot{
            .files                              = _files,
            .directories                        = _directories,
            .commands                           = _commands,
            .deletedPaths                       = _deletedPaths,
            .lastUploadReceivedBytes            = _lastUploadReceivedBytes,
            .lastUploadRetainedBytes            = _lastUploadRetainedBytes,
            .lastDownloadSentBytes              = _lastDownloadSentBytes,
            .beforeRenameInjectionCount         = _beforeRenameInjectionCount,
            .afterRenameInjectionCount          = _afterRenameInjectionCount,
            .beforeListInjectionCount           = _beforeListInjectionCount,
            .afterDirectoryCreateInjectionCount = _afterDirectoryCreateInjectionCount,
            .rejectedDeleteCount                = _rejectedDeleteCount,
            .lostMutationReplyCount             = _lostMutationReplyCount,
            .rejectedDirectoryAccessCount       = _rejectedDirectoryAccessCount,
            .lateListingReachedCount            = _lateListingReachedCount,
            .lateListingReleasedCount           = _lateListingReleasedCount,
            .lateListingTimeoutCount            = _lateListingTimeoutCount,
            .passiveListenerCreations           = _passiveListenerCreations,
            .passiveListenerReuses              = _passiveListenerReuses,
            .passivePendingRetirements          = _passivePendingRetirements,
            .idleControlClosures                = _idleControlClosures,
            .dataTransfers                      = _dataTransfers,
            .listPayloadBuildUs                 = _listPayloadBuildUs,
            .listPayloadBytes                   = _listPayloadBytes,
            .serverHr                           = _serverHr,
        };
    }

private:
    // Fixture protocol state, not a production connection pool. Reuse a session's
    // listening port only after its advertised data connection was consumed;
    // every LIST/RETR/STOR still owns a fresh accepted data socket.
    struct PassiveListener final
    {
        PassiveListener()                                  = default;
        PassiveListener(const PassiveListener&)            = delete;
        PassiveListener& operator=(const PassiveListener&) = delete;
        PassiveListener(PassiveListener&&)                 = delete;
        PassiveListener& operator=(PassiveListener&&)      = delete;

        wil::unique_socket socket;
        unsigned short port = 0u;
        bool pending        = false;
        bool used           = false;
    };

    void SetServerFailure(HRESULT hr) noexcept
    {
        std::scoped_lock lock(_stateMutex);
        if (SUCCEEDED(_serverHr))
        {
            _serverHr = hr;
        }
    }

    void RecordCommand(std::string verb, std::string argument)
    {
        std::scoped_lock lock(_stateMutex);
        _commands.push_back(CommandRecord{.verb = std::move(verb), .argument = std::move(argument)});
    }

    void AddDirectoryWithParentsLocked(std::string_view path)
    {
        _directories.insert("/");
        if (path.empty() || path == "/")
        {
            return;
        }

        const std::string normalized = NormalizeFtpPath(path, "/");
        size_t separator             = 1u;
        for (;;)
        {
            separator = normalized.find('/', separator);
            _directories.insert(separator == std::string::npos ? normalized : normalized.substr(0u, separator));
            if (separator == std::string::npos)
            {
                break;
            }
            ++separator;
        }
    }

    void InjectFileBeforeRename(std::string_view path)
    {
        std::scoped_lock lock(_stateMutex);
        const auto found = _filesBeforeRename.find(std::string(path));
        if (found == _filesBeforeRename.end())
        {
            return;
        }

        AddDirectoryWithParentsLocked(FtpFixtureParentView(found->first));
        _files[found->first] = std::move(found->second);
        _filesBeforeRename.erase(found);
        ++_beforeRenameInjectionCount;
    }

    void InjectFileAfterRename(std::string_view path)
    {
        std::scoped_lock lock(_stateMutex);
        const auto found = _filesAfterRename.find(std::string(path));
        if (found == _filesAfterRename.end())
        {
            return;
        }

        AddDirectoryWithParentsLocked(FtpFixtureParentView(found->first));
        _files[found->first] = std::move(found->second);
        _filesAfterRename.erase(found);
        ++_afterRenameInjectionCount;
    }

    void InjectFileBeforeDirectoryList(std::string_view directory)
    {
        std::scoped_lock lock(_stateMutex);
        const auto found = _filesBeforeDirectoryList.find(std::string(directory));
        if (found == _filesBeforeDirectoryList.end())
        {
            return;
        }

        AddDirectoryWithParentsLocked(FtpFixtureParentView(found->second.path));
        _files[found->second.path] = std::move(found->second.bytes);
        _filesBeforeDirectoryList.erase(found);
        ++_beforeListInjectionCount;
    }

    void InjectFileAfterDirectoryCreate(std::string_view directory)
    {
        std::scoped_lock lock(_stateMutex);
        const auto found = _filesAfterDirectoryCreate.find(std::string(directory));
        if (found == _filesAfterDirectoryCreate.end())
        {
            return;
        }

        AddDirectoryWithParentsLocked(FtpFixtureParentView(found->second.path));
        _files[found->second.path] = std::move(found->second.bytes);
        _filesAfterDirectoryCreate.erase(found);
        ++_afterDirectoryCreateInjectionCount;
    }

    [[nodiscard]] bool ShouldRejectDelete(std::string_view path)
    {
        std::scoped_lock lock(_stateMutex);
        const bool exactFailure = _deleteFailures.contains(std::string(path));
        const bool markerFailure = std::ranges::any_of(
            _deleteFailureMarkers, [path](const std::string& marker) noexcept { return path.find(marker) != std::string_view::npos; });
        if (! exactFailure && ! markerFailure)
        {
            return false;
        }
        ++_rejectedDeleteCount;
        return true;
    }

    [[nodiscard]] bool CanEnterDirectory(std::string_view path)
    {
        std::scoped_lock lock(_stateMutex);
        if (_deniedDirectories.contains(std::string(path)))
        {
            ++_rejectedDirectoryAccessCount;
            return false;
        }
        return _directories.contains(std::string(path));
    }

    [[nodiscard]] bool RemoveDirectory(std::string_view path)
    {
        const std::string normalized(path);
        if (normalized == "/")
        {
            return false;
        }

        std::scoped_lock lock(_stateMutex);
        if (! _directories.contains(normalized))
        {
            return false;
        }
        for (const auto& [filePath, bytes] : _files)
        {
            static_cast<void>(bytes);
            if (FtpFixtureParentView(filePath) == normalized)
            {
                return false;
            }
        }
        for (const std::string& directory : _directories)
        {
            if (directory != normalized && FtpFixtureParentView(directory) == normalized)
            {
                return false;
            }
        }
        return _directories.erase(normalized) == 1u;
    }

public:
    // This class is translation-unit-local test infrastructure. Byte oracles
    // exercise the exact serializer used by HandleList, not a second encoder.
    [[nodiscard]] std::string BuildDirectoryPayload(std::string_view directory, bool namesOnly) const
    {
        std::scoped_lock lock(_stateMutex);
        const auto custom = _directoryPayloads.find(std::string(directory));
        if (custom != _directoryPayloads.end())
        {
            return custom->second;
        }
        std::string payload;
        for (const std::string& childDirectory : _directories)
        {
            if (childDirectory == "/" || FtpFixtureParentView(childDirectory) != directory)
            {
                continue;
            }

            const std::string_view leaf = FtpFixtureLeafView(childDirectory);
            if (namesOnly)
            {
                payload.append(leaf);
                payload.append("\r\n");
            }
            else
            {
                std::format_to(std::back_inserter(payload), "drwxr-xr-x 1 owner group 0 Jan 01 2026 {}\r\n", leaf);
            }
        }
        for (const auto& [path, bytes] : _files)
        {
            if (FtpFixtureParentView(path) != directory)
            {
                continue;
            }
            const std::string_view leaf = FtpFixtureLeafView(path);
            if (leaf.empty())
            {
                continue;
            }

            if (namesOnly)
            {
                payload.append(leaf);
                payload.append("\r\n");
            }
            else if (_listIncludesSize)
            {
                const auto listedSize = _listedSizes.find(path);
                const size_t size     = listedSize == _listedSizes.end() ? bytes.size() : listedSize->second;
                std::format_to(std::back_inserter(payload), "-rw-r--r-- 1 owner group {} Jan 01 2026 {}\r\n", size, leaf);
            }
            else
            {
                // Keep the record structurally parseable so a mixed directory
                // listing cannot discard this file after parsing a directory
                // entry. The nonnumeric size token intentionally yields
                // sizeKnown=false in the production Unix LIST parser.
                std::format_to(std::back_inserter(payload), "-rw-r--r-- 1 owner group ? Jan 01 2026 {}\r\n", leaf);
            }
        }
        return payload;
    }

private:
    [[nodiscard]] bool TryReadFile(std::string_view path, std::vector<uint8_t>& bytesOut) const
    {
        std::scoped_lock lock(_stateMutex);
        const auto found = _files.find(std::string(path));
        if (found == _files.end())
        {
            return false;
        }
        bytesOut = found->second;
        return true;
    }

    [[nodiscard]] bool TryGetFileSize(std::string_view path, size_t& sizeOut) const noexcept
    {
        std::scoped_lock lock(_stateMutex);
        const auto found = _files.find(std::string(path));
        if (found == _files.end())
        {
            return false;
        }
        sizeOut = found->second.size();
        return true;
    }

    // Fixture-only atomic interleaving: commit, plant a different source generation, then
    // consume the one-shot fault before dropping the reply. No guessed-path cleanup is involved.
    [[nodiscard]] bool ReplaceAfterCommittedMutationLocked(std::string_view verb, std::string_view sourcePath)
    {
        if (_lostMutationReplyVerb != verb || _lostMutationReplyPath != sourcePath)
        {
            return false;
        }
        _files[_lostMutationReplyPath] = std::move(_lostMutationReplyReplacement);
        _lostMutationReplyVerb.clear();
        ++_lostMutationReplyCount;
        return true;
    }

    [[nodiscard]] bool DeleteFile(std::string_view path, bool& loseReply)
    {
        loseReply = false;
        std::scoped_lock lock(_stateMutex);
        const size_t erased = _files.erase(std::string(path));
        if (erased != 0u)
        {
            _deletedPaths.emplace_back(path);
            _lateListingChanged.notify_all();
            loseReply = ReplaceAfterCommittedMutationLocked("DELE", path);
            return true;
        }
        return false;
    }

    [[nodiscard]] bool RenameEntry(std::string_view sourcePath, std::string_view destinationPath, bool& loseReply)
    {
        loseReply = false;
        std::scoped_lock lock(_stateMutex);
        const auto found = _files.find(std::string(sourcePath));
        if (found != _files.end())
        {
            if (_directories.contains(std::string(destinationPath)))
            {
                return false;
            }
            std::vector<uint8_t> bytes = std::move(found->second);
            _files.erase(found);
            _files[std::string(destinationPath)] = std::move(bytes);
            loseReply                            = ReplaceAfterCommittedMutationLocked("RNTO", sourcePath);
            return true;
        }

        // Fixture-only server rename: move the exact directory subtree atomically
        // under the endpoint lock. Existing targets are refused, not recursively merged.
        const std::string source(sourcePath);
        const std::string destination(destinationPath);
        const std::string sourcePrefix = source + '/';
        if (source == _refusedDirectoryRename)
        {
            return false;
        }
        if (source == "/" || ! _directories.contains(source) || destination.starts_with(sourcePrefix) || _directories.contains(destination) ||
            _files.contains(destination) || ! _directories.contains(std::string(FtpFixtureParentView(destination))))
        {
            return false;
        }

        for (auto entry = _files.lower_bound(sourcePrefix); entry != _files.end() && entry->first.starts_with(sourcePrefix);)
        {
            auto node  = _files.extract(entry++);
            node.key() = destination + node.key().substr(source.size());
            _files.insert(std::move(node));
        }
        for (auto entry = _directories.lower_bound(sourcePrefix); entry != _directories.end() && entry->starts_with(sourcePrefix);)
        {
            auto node    = _directories.extract(entry++);
            node.value() = destination + node.value().substr(source.size());
            _directories.insert(std::move(node));
        }
        auto root    = _directories.extract(source);
        root.value() = destination;
        _directories.insert(std::move(root));
        return true;
    }

    void StoreUpload(std::string path, std::vector<uint8_t> received)
    {
        size_t retainedSize = received.size();
        if (_retention == UploadRetention::StrictPrefix && ! received.empty())
        {
            retainedSize = (std::max)(size_t{1u}, received.size() / 2u);
            if (retainedSize >= received.size())
            {
                retainedSize = received.size() - 1u;
            }
        }

        std::vector<uint8_t> retained(received.begin(), received.begin() + static_cast<std::ptrdiff_t>(retainedSize));
        std::scoped_lock lock(_stateMutex);
        _lastUploadReceivedBytes = received.size();
        _lastUploadRetainedBytes = retained.size();
        _files[std::move(path)]  = std::move(retained);
    }

    [[nodiscard]] HRESULT AcceptDataConnection(const wil::unique_socket& passiveListener, std::stop_token stopToken, wil::unique_socket& dataOut) noexcept
    {
        dataOut.reset();
        if (! passiveListener)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        const HRESULT waitHr = WaitForReadable(passiveListener.get(), stopToken, std::chrono::steady_clock::now() + kDataOperationTimeout);
        if (FAILED(waitHr))
        {
            return waitHr;
        }

        wil::unique_socket data(accept(passiveListener.get(), nullptr, nullptr));
        if (! data)
        {
            return SocketErrorToHResult();
        }
        const HRESULT timeoutHr = ConfigureSocketTimeouts(data.get());
        if (FAILED(timeoutHr))
        {
            return timeoutHr;
        }
        dataOut = std::move(data);
        return S_OK;
    }

    [[nodiscard]] HRESULT ReceiveData(SOCKET socketValue, std::stop_token stopToken, std::vector<uint8_t>& bytesOut)
    {
        bytesOut.clear();
        auto deadline = std::chrono::steady_clock::now() + kDataOperationTimeout;
        for (;;)
        {
            const HRESULT waitHr = WaitForReadable(socketValue, stopToken, deadline);
            if (FAILED(waitHr))
            {
                return waitHr;
            }

            std::array<char, 4096u> buffer{};
            const int received = recv(socketValue, buffer.data(), static_cast<int>(buffer.size()), 0);
            if (received == 0)
            {
                return S_OK;
            }
            if (received == SOCKET_ERROR)
            {
                const int error = WSAGetLastError();
                if (error == WSAETIMEDOUT || error == WSAEWOULDBLOCK)
                {
                    continue;
                }
                return HRESULT_FROM_WIN32(static_cast<unsigned long>(error));
            }

            const size_t receivedSize = static_cast<size_t>(received);
            if (bytesOut.size() > kMaximumDataTransferBytes - receivedSize)
            {
                return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
            }
            const auto* first = reinterpret_cast<const uint8_t*>(buffer.data());
            bytesOut.insert(bytesOut.end(), first, first + receivedSize);
            deadline = std::chrono::steady_clock::now() + kDataOperationTimeout;
        }
    }

    [[nodiscard]] HRESULT BeginDataTransfer(
        SOCKET control, PassiveListener& passiveListener, DataTransferKind kind, std::stop_token stopToken, wil::unique_socket& dataOut) noexcept
    {
        if (! passiveListener.socket || ! std::exchange(passiveListener.pending, false))
        {
            static_cast<void>(SendAll(control, "425 Use EPSV or PASV first.\r\n"));
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        HRESULT hr = SendAll(control, "150 Opening loopback data connection.\r\n");
        if (FAILED(hr))
        {
            return hr;
        }
        hr = AcceptDataConnection(passiveListener.socket, stopToken, dataOut);
        if (FAILED(hr))
        {
            static_cast<void>(SendAll(control, "425 Cannot open data connection.\r\n"));
        }
        else
        {
            std::scoped_lock lock(_stateMutex);
            DataTransferCounts& counts = _dataTransfers[static_cast<size_t>(kind)];
            ++(passiveListener.used ? counts.reuse : counts.firstUse);
            passiveListener.used = true;
        }
        return hr;
    }

    [[nodiscard]] HRESULT HandleList(SOCKET control, PassiveListener& passiveListener, std::stop_token stopToken, std::string_view directory, bool namesOnly)
    {
        wil::unique_socket data;
        HRESULT hr = BeginDataTransfer(control, passiveListener, namesOnly ? DataTransferKind::NamesOnly : DataTransferKind::List, stopToken, data);
        if (FAILED(hr))
        {
            return hr;
        }

        const auto buildStarted   = std::chrono::steady_clock::now();
        const std::string payload = BuildDirectoryPayload(directory, namesOnly);
        const uint64_t buildUs    = Debug::Perf::ElapsedUs(buildStarted);
        bool failFinalReply       = false;
        {
            std::scoped_lock lock(_stateMutex);
            _listPayloadBuildUs += buildUs;
            _listPayloadBytes += payload.size();
            failFinalReply = _failListFinalReply;
        }
        hr = SendAll(data.get(), payload);
        data.reset();
        if (FAILED(hr))
        {
            static_cast<void>(SendAll(control, "426 Data transfer aborted.\r\n"));
            return hr;
        }
        return SendAll(control, failFinalReply ? "426 Directory transfer failed.\r\n" : "226 Directory transfer complete.\r\n");
    }

    [[nodiscard]] HRESULT HandleRetrieve(
        SOCKET control, PassiveListener& passiveListener, std::stop_token stopToken, std::string_view path, uint64_t restartOffset)
    {
        std::vector<uint8_t> bytes;
        if (! TryReadFile(path, bytes))
        {
            return SendAll(control, "550 File not found.\r\n");
        }

        wil::unique_socket data;
        HRESULT hr = BeginDataTransfer(control, passiveListener, DataTransferKind::Retrieve, stopToken, data);
        if (FAILED(hr))
        {
            return hr;
        }

        const size_t offset = static_cast<size_t>((std::min)(restartOffset, static_cast<uint64_t>(bytes.size())));
        std::vector<uint8_t> delivered(bytes.begin() + static_cast<std::ptrdiff_t>(offset), bytes.end());
        DownloadDelivery delivery = DownloadDelivery::Complete;
        {
            std::scoped_lock lock(_stateMutex);
            delivery = _downloadDelivery;
        }
        if (delivery == DownloadDelivery::StrictPrefix && ! delivered.empty())
        {
            size_t retainedSize = (std::max)(size_t{1u}, delivered.size() / 2u);
            if (retainedSize >= delivered.size())
            {
                retainedSize = delivered.size() - 1u;
            }
            delivered.resize(retainedSize);
        }
        else if (delivery == DownloadDelivery::Overlong)
        {
            delivered.push_back(0xa5u);
        }

        if (! delivered.empty())
        {
            const std::string_view payload(reinterpret_cast<const char*>(delivered.data()), delivered.size());
            hr = SendAll(data.get(), payload);
        }
        else
        {
            hr = S_OK;
        }
        data.reset();
        if (FAILED(hr))
        {
            static_cast<void>(SendAll(control, "426 Data transfer aborted.\r\n"));
            return hr;
        }
        {
            std::scoped_lock lock(_stateMutex);
            _lastDownloadSentBytes = delivered.size();
            if (_fileAfterNextRetrievePending)
            {
                AddDirectoryWithParentsLocked(FtpFixtureParentView(_fileAfterNextRetrievePath));
                _files[_fileAfterNextRetrievePath] = std::move(_fileAfterNextRetrieveBytes);
                _fileAfterNextRetrievePath.clear();
                _fileAfterNextRetrievePending = false;
            }
        }
        return SendAll(control, "226 Download complete.\r\n");
    }

    [[nodiscard]] HRESULT HandleStore(SOCKET control, PassiveListener& passiveListener, std::stop_token stopToken, std::string path)
    {
        wil::unique_socket data;
        HRESULT hr = BeginDataTransfer(control, passiveListener, DataTransferKind::Store, stopToken, data);
        if (FAILED(hr))
        {
            return hr;
        }

        std::vector<uint8_t> received;
        hr = ReceiveData(data.get(), stopToken, received);
        data.reset();
        if (FAILED(hr))
        {
            static_cast<void>(SendAll(control, "426 Upload aborted.\r\n"));
            return hr;
        }

        // Fault modes consume the complete Curl stream so the operation-owned staging path exists for cleanup validation.
        StoreUpload(std::move(path), std::move(received));
        if (_retention == UploadRetention::FailAfterStore)
        {
            return SendAll(control, "451 Upload rejected after deterministic storage.\r\n");
        }
        return SendAll(control, "226 Upload complete.\r\n");
    }

    [[nodiscard]] HRESULT EnterPassiveMode(SOCKET control, PassiveListener& passiveListener, bool extended)
    {
        if (passiveListener.pending)
        {
            // A rejected RETR/LIST or repeated EPSV may leave an old data
            // connection queued. Retire its listener instead of accepting it
            // as the next command's data channel.
            passiveListener.socket.reset();
            passiveListener.pending = false;
            std::scoped_lock lock(_stateMutex);
            ++_passivePendingRetirements;
        }
        if (! passiveListener.socket)
        {
            const HRESULT hr = CreateLoopbackEndpoint(passiveListener.socket, passiveListener.port);
            if (FAILED(hr))
            {
                static_cast<void>(SendAll(control, "425 Cannot enter passive mode.\r\n"));
                return hr;
            }
            std::scoped_lock lock(_stateMutex);
            ++_passiveListenerCreations;
            passiveListener.used = false;
        }
        else
        {
            std::scoped_lock lock(_stateMutex);
            ++_passiveListenerReuses;
        }
        passiveListener.pending   = true;
        const unsigned short port = passiveListener.port;

        if (extended)
        {
            return SendAll(control, std::format("229 Entering Extended Passive Mode (|||{}|).\r\n", port));
        }

        const unsigned int high = static_cast<unsigned int>(port) / 256u;
        const unsigned int low  = static_cast<unsigned int>(port) % 256u;
        return SendAll(control, std::format("227 Entering Passive Mode (127,0,0,1,{},{}).\r\n", high, low));
    }

    [[nodiscard]] HRESULT HandleClient(SOCKET control, std::stop_token stopToken)
    {
        HRESULT hr = ConfigureSocketTimeouts(control);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = SendAll(control, "220 RedSalamander deterministic FTP fixture ready.\r\n");
        if (FAILED(hr))
        {
            return hr;
        }

        std::string pending;
        std::string currentDirectory = "/";
        std::string renameFrom;
        uint64_t restartOffset = 0u;
        PassiveListener passiveListener;

        while (! stopToken.stop_requested())
        {
            std::string line;
            hr = ReadControlLine(control, stopToken, pending, line);
            if (FAILED(hr))
            {
                if (hr == HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT) && pending.empty())
                {
                    // A quiet pooled/paused control connection may expire like
                    // a real FTP session. Partial commands and active data waits
                    // still fail; exact scenario bytes/status remain mandatory.
                    {
                        std::scoped_lock lock(_stateMutex);
                        ++_idleControlClosures;
                    }
                    static_cast<void>(SendAll(control, "421 Idle control connection closed.\r\n"));
                    return HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED);
                }
                return hr;
            }

            const size_t separator     = line.find(' ');
            const std::string verb     = UpperAscii(separator == std::string::npos ? std::string_view(line) : std::string_view(line).substr(0u, separator));
            const std::string argument = separator == std::string::npos ? std::string{} : line.substr(separator + 1u);
            RecordCommand(verb, argument);

            {
                unsigned int stallMs = 0u;
                {
                    std::scoped_lock lock(_stateMutex);
                    if (! _stallVerb.empty() && verb == _stallVerb)
                    {
                        stallMs = _stallMs;
                    }
                }
                const ULONGLONG stallUntil = GetTickCount64() + stallMs;
                while (stallMs != 0u && ! stopToken.stop_requested() && GetTickCount64() < stallUntil)
                {
                    Sleep(20u);
                }
            }

            if (verb == "LIST" || verb == "NLST")
            {
                const std::string directory = argument.empty() ? currentDirectory : NormalizeFtpPath(argument, currentDirectory);
                bool refuse                 = false;
                {
                    std::unique_lock lock(_stateMutex);
                    if (! _lateListingDirectory.empty() && directory == _lateListingDirectory)
                    {
                        ++_lateListingReachedCount;
                        _lateListingChanged.notify_all();
                        const bool priorDelete = _lateListingChanged.wait_for(lock, stopToken, 5s, [this]() noexcept {
                            return std::ranges::find(_deletedPaths, _lateListingDeletedPath) != _deletedPaths.end();
                        });
                        if (stopToken.stop_requested())
                        {
                            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        }
                        if (priorDelete)
                        {
                            ++_lateListingReleasedCount;
                        }
                        else
                        {
                            ++_lateListingTimeoutCount;
                        }
                        _lateListingChanged.notify_all();
                        if (priorDelete && _lateListingReply == LateListingReply::WaitForCancel)
                        {
                            // Stop() owns this unblock. A provider that only returns after
                            // fixture shutdown must fail the cooperative-cancel assertion.
                            static_cast<void>(_lateListingChanged.wait(lock, stopToken, []() noexcept { return false; }));
                            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        }
                        refuse = ! priorDelete || _lateListingReply == LateListingReply::Refuse;
                    }
                }
                if (refuse)
                {
                    hr = SendAll(control, "550 Late directory listing refused.\r\n");
                    if (FAILED(hr))
                    {
                        return hr;
                    }
                    continue;
                }
                // Listing data can be backpressured by a bounded client cursor.
                // Snapshot construction owns _stateMutex, but network delivery must
                // not block DELE on another session behind _commandMutex.
                InjectFileBeforeDirectoryList(directory);
                hr = HandleList(control, passiveListener, stopToken, directory, verb == "NLST");
                if (FAILED(hr))
                {
                    return hr;
                }
                continue;
            }

            // One command at a time across all sessions keeps the in-memory endpoint deterministic.
            std::scoped_lock commandLock(_commandMutex);

            if (verb == "USER")
            {
                hr = SendAll(control, "331 Password accepted by fixture.\r\n");
            }
            else if (verb == "PASS")
            {
                hr = SendAll(control, "230 Logged in.\r\n");
            }
            else if (verb == "SYST")
            {
                hr = SendAll(control, "215 UNIX Type: L8.\r\n");
            }
            else if (verb == "FEAT")
            {
                hr =
                    SendAll(control,
                            _sizeCommandSupported ? "211-Features\r\n EPSV\r\n SIZE\r\n UTF8\r\n211 End\r\n" : "211-Features\r\n EPSV\r\n UTF8\r\n211 End\r\n");
            }
            else if (verb == "OPTS" || verb == "CLNT" || verb == "NOOP")
            {
                hr = SendAll(control, "200 Command accepted.\r\n");
            }
            else if (verb == "PWD" || verb == "XPWD")
            {
                hr = SendAll(control, std::format("257 \"{}\" is the current directory.\r\n", currentDirectory));
            }
            else if (verb == "CWD")
            {
                const std::string requestedDirectory = NormalizeFtpPath(argument, currentDirectory);
                if (CanEnterDirectory(requestedDirectory))
                {
                    currentDirectory = requestedDirectory;
                    hr               = SendAll(control, "250 Directory changed.\r\n");
                }
                else
                {
                    hr = SendAll(control, "550 Directory not found.\r\n");
                }
            }
            else if (verb == "CDUP")
            {
                currentDirectory = FtpFixtureParentView(currentDirectory);
                hr               = SendAll(control, "250 Directory changed.\r\n");
            }
            else if (verb == "TYPE")
            {
                hr = SendAll(control, "200 Transfer type set.\r\n");
            }
            else if (verb == "EPSV")
            {
                hr = EnterPassiveMode(control, passiveListener, true);
            }
            else if (verb == "PASV")
            {
                hr = EnterPassiveMode(control, passiveListener, false);
            }
            else if (verb == "SIZE")
            {
                const std::string path = NormalizeFtpPath(argument, currentDirectory);
                size_t size            = 0u;
                if (_sizeAuthenticationFailure.load(std::memory_order_acquire))
                {
                    hr = SendAll(control, "530 Authentication required.\r\n");
                }
                else if (_sizeMissingFailure.load(std::memory_order_acquire))
                {
                    hr = SendAll(control, "550 File disappeared before SIZE.\r\n");
                }
                else if (! _sizeCommandSupported)
                {
                    hr = SendAll(control, "502 SIZE is not supported.\r\n");
                }
                else
                {
                    hr = TryGetFileSize(path, size) ? SendAll(control, std::format("213 {}\r\n", size))
                                                    : SendAll(control, "550 File not found.\r\n");
                }
            }
            else if (verb == "MDTM")
            {
                const std::string path = NormalizeFtpPath(argument, currentDirectory);
                size_t ignored         = 0u;
                hr = TryGetFileSize(path, ignored) ? SendAll(control, "213 20260721000000\r\n") : SendAll(control, "550 File not found.\r\n");
            }
            else if (verb == "REST")
            {
                uint64_t parsedOffset = 0u;
                bool validOffset      = ! argument.empty();
                for (const char ch : argument)
                {
                    if (ch < '0' || ch > '9')
                    {
                        validOffset = false;
                        break;
                    }
                    const uint64_t digit = static_cast<uint64_t>(ch - '0');
                    if (parsedOffset > ((std::numeric_limits<uint64_t>::max)() - digit) / 10u)
                    {
                        validOffset = false;
                        break;
                    }
                    parsedOffset = parsedOffset * 10u + digit;
                }
                if (validOffset)
                {
                    restartOffset = parsedOffset;
                    hr            = SendAll(control, "350 Restart position accepted.\r\n");
                }
                else
                {
                    hr = SendAll(control, "501 Invalid restart position.\r\n");
                }
            }
            else if (verb == "RETR")
            {
                hr = HandleRetrieve(control, passiveListener, stopToken, NormalizeFtpPath(argument, currentDirectory), restartOffset);
                restartOffset = 0u;
            }
            else if (verb == "STOR")
            {
                hr = HandleStore(control, passiveListener, stopToken, NormalizeFtpPath(argument, currentDirectory));
            }
            else if (verb == "DELE")
            {
                const std::string path = NormalizeFtpPath(argument, currentDirectory);
                bool loseReply = false;
                hr = ShouldRejectDelete(path)
                         ? SendAll(control, "550 Delete rejected by deterministic race fixture.\r\n")
                         : (DeleteFile(path, loseReply) ? (loseReply ? HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED) : SendAll(control, "250 File deleted.\r\n"))
                                                        : SendAll(control, "550 File not found.\r\n"));
            }
            else if (verb == "RNFR")
            {
                renameFrom  = NormalizeFtpPath(argument, currentDirectory);
                bool exists = false;
                {
                    std::scoped_lock lock(_stateMutex);
                    exists = _files.contains(renameFrom) || _directories.contains(renameFrom);
                }
                if (exists)
                {
                    hr = SendAll(control, "350 Ready for destination name.\r\n");
                }
                else
                {
                    renameFrom.clear();
                    hr = SendAll(control, "550 File not found.\r\n");
                }
            }
            else if (verb == "RNTO")
            {
                const std::string destination = NormalizeFtpPath(argument, currentDirectory);
                InjectFileBeforeRename(destination);
                const bool rejectWriterPromotion = _failWriterPromotion && renameFrom.find(kWriterSiblingMarker) != std::string::npos;
                bool loseReply                   = false;
                const bool renamed               = ! rejectWriterPromotion && ! renameFrom.empty() && RenameEntry(renameFrom, destination, loseReply);
                if (renamed)
                {
                    InjectFileAfterRename(destination);
                }
                renameFrom.clear();
                hr = renamed ? (loseReply ? HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED) : SendAll(control, "250 Rename complete.\r\n"))
                             : SendAll(control, "550 Rename rejected.\r\n");
            }
            else if (verb == "MKD")
            {
                const std::string path = NormalizeFtpPath(argument, currentDirectory);
                {
                    std::scoped_lock lock(_stateMutex);
                    AddDirectoryWithParentsLocked(path);
                }
                InjectFileAfterDirectoryCreate(path);
                hr = SendAll(control, "257 Directory created.\r\n");
            }
            else if (verb == "RMD")
            {
                const std::string path = NormalizeFtpPath(argument, currentDirectory);
                hr = RemoveDirectory(path) ? SendAll(control, "250 Directory removed.\r\n") : SendAll(control, "550 Directory not empty or absent.\r\n");
            }
            else if (verb == "ABOR")
            {
                passiveListener.socket.reset();
                passiveListener.pending = false;
                passiveListener.port    = 0u;
                hr                      = SendAll(control, "226 Abort acknowledged.\r\n");
            }
            else if (verb == "QUIT")
            {
                return SendAll(control, "221 Goodbye.\r\n");
            }
            else
            {
                hr = SendAll(control, "502 Command not implemented by fixture.\r\n");
            }

            if (FAILED(hr))
            {
                return hr;
            }
        }

        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    void ServerMain(std::stop_token stopToken)
    {
        while (! stopToken.stop_requested())
        {
            const HRESULT waitHr = WaitForReadable(_listener.get(), stopToken, std::chrono::steady_clock::now() + kSocketPollInterval);
            if (waitHr == HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT))
            {
                continue;
            }
            if (waitHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return;
            }
            if (FAILED(waitHr))
            {
                SetServerFailure(waitHr);
                return;
            }

            wil::unique_socket control(accept(_listener.get(), nullptr, nullptr));
            if (! control)
            {
                SetServerFailure(SocketErrorToHResult());
                return;
            }

            // Real servers serve sessions concurrently, and the host bridge keeps a streaming
            // transfer session open while it probes on other connections. Each session runs on
            // its own thread; command processing stays serialized (see HandleClient).
            std::scoped_lock lock(_clientsMutex);
            std::erase_if(_clients, [](const std::jthread& client) noexcept { return ! client.joinable(); });
            _clients.emplace_back([this, ownedControl = std::move(control)](std::stop_token clientStopToken) noexcept
            {
                HRESULT clientHr = E_FAIL;
                try
                {
                    clientHr = HandleClient(ownedControl.get(), clientStopToken);
                }
                catch (const std::bad_alloc&)
                {
                    std::terminate();
                }
                catch (const std::exception&)
                {
                    // Mandatory thread boundary: a fixture exception must become a visible test failure.
                    Debug::Error(L"FileSystemCurl fake FTP client failed after std::exception.");
                    clientHr = E_FAIL;
                }
                // An early-exit streaming consumer can reset an unfinished LIST
                // session. Winsock can report either reset or local connection
                // abort while that session drains. Like EOF/cancel these end the
                // session; each scenario still requires exact bytes, operation
                // status and mutation truth, so a premature disconnect cannot pass.
                if (clientHr == HRESULT_FROM_WIN32(WSAECONNABORTED))
                {
                    Debug::Perf::Emit(L"FileOps.Curl.FakeFtp.PeerDisconnect",
                                      L"connection-aborted;scenario-validates-state",
                                      0u,
                                      static_cast<unsigned long>(clientHr),
                                      0u,
                                      S_OK);
                }
                if (FAILED(clientHr) && clientHr != HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED) && clientHr != HRESULT_FROM_WIN32(WSAECONNRESET) &&
                    clientHr != HRESULT_FROM_WIN32(WSAECONNABORTED) && clientHr != HRESULT_FROM_WIN32(ERROR_CANCELLED))
                {
                    SetServerFailure(clientHr);
                }
            });
        }
    }

    UploadRetention _retention = UploadRetention::Complete;
    bool _sizeCommandSupported = true;
    std::string _stallVerb;
    unsigned int _stallMs = 0u;
    std::atomic_bool _sizeAuthenticationFailure{false};
    std::atomic_bool _sizeMissingFailure{false};
    bool _listIncludesSize             = true;
    bool _failWriterPromotion          = false;
    DownloadDelivery _downloadDelivery = DownloadDelivery::Complete;
    unsigned short _port       = 0u;
    wil::unique_socket _listener;
    std::mutex _clientsMutex;
    std::vector<std::jthread> _clients;
    std::mutex _commandMutex;

    mutable std::mutex _stateMutex;
    std::condition_variable_any _lateListingChanged;
    std::string _lateListingDirectory;
    std::string _lateListingDeletedPath;
    LateListingReply _lateListingReply = LateListingReply::Complete;
    size_t _lateListingReachedCount = 0u;
    size_t _lateListingReleasedCount   = 0u;
    size_t _lateListingTimeoutCount    = 0u;
    struct PendingFileInjection final
    {
        std::string path;
        std::vector<uint8_t> bytes;
    };
    std::map<std::string, std::vector<uint8_t>> _files;
    std::string _refusedDirectoryRename;
    std::map<std::string, std::string> _directoryPayloads;
    bool _failListFinalReply = false;
    std::map<std::string, size_t> _listedSizes;
    std::map<std::string, std::vector<uint8_t>> _filesBeforeRename;
    std::map<std::string, std::vector<uint8_t>> _filesAfterRename;
    std::string _lostMutationReplyVerb;
    std::string _lostMutationReplyPath;
    std::vector<uint8_t> _lostMutationReplyReplacement;
    std::map<std::string, PendingFileInjection> _filesBeforeDirectoryList;
    std::map<std::string, PendingFileInjection> _filesAfterDirectoryCreate;
    std::string _fileAfterNextRetrievePath;
    std::vector<uint8_t> _fileAfterNextRetrieveBytes;
    bool _fileAfterNextRetrievePending = false;
    std::set<std::string> _deleteFailures;
    std::set<std::string> _deniedDirectories;
    std::set<std::string> _deleteFailureMarkers;
    std::set<std::string> _directories{"/"};
    std::vector<CommandRecord> _commands;
    std::vector<std::string> _deletedPaths;
    size_t _lastUploadReceivedBytes = 0u;
    size_t _lastUploadRetainedBytes = 0u;
    size_t _lastDownloadSentBytes   = 0u;
    size_t _beforeRenameInjectionCount = 0u;
    size_t _afterRenameInjectionCount  = 0u;
    size_t _beforeListInjectionCount   = 0u;
    size_t _afterDirectoryCreateInjectionCount = 0u;
    size_t _rejectedDeleteCount        = 0u;
    size_t _lostMutationReplyCount     = 0u;
    size_t _rejectedDirectoryAccessCount = 0u;
    size_t _passiveListenerCreations     = 0u;
    size_t _passiveListenerReuses        = 0u;
    size_t _passivePendingRetirements    = 0u;
    size_t _idleControlClosures          = 0u;
    std::array<DataTransferCounts, static_cast<size_t>(DataTransferKind::Count)> _dataTransfers{};
    uint64_t _listPayloadBuildUs    = 0u;
    uint64_t _listPayloadBytes      = 0u;
    HRESULT _serverHr               = S_OK;

    std::jthread _thread;
};

struct ScenarioResult final
{
    HRESULT setupHr    = E_FAIL;
    HRESULT moveHr     = E_FAIL;
    uint64_t elapsedMs = 0u;
    EndpointSnapshot source;
    EndpointSnapshot destination;
};

struct WriterLifetimeResult final
{
    HRESULT setupHr       = E_FAIL;
    HRESULT createHr      = E_FAIL;
    HRESULT writeHr       = E_FAIL;
    HRESULT commitHr      = E_FAIL;
    unsigned long written = 0u;
    uint64_t elapsedMs    = 0u;
    EndpointSnapshot destination;
};

struct ReaderScenarioResult final
{
    HRESULT setupHr      = E_FAIL;
    HRESULT createHr     = E_FAIL;
    HRESULT getSizeHr    = E_FAIL;
    HRESULT seekHr       = S_OK;
    HRESULT firstReadHr  = E_FAIL;
    HRESULT secondReadHr = E_FAIL;
    HRESULT restartSeekHr = E_FAIL;
    HRESULT restartReadHr = E_FAIL;
    HRESULT restartEofHr  = E_FAIL;
    uint64_t committedSizeBytes = 0u;
    uint64_t seekPosition       = 0u;
    unsigned long firstBytesRead  = 0u;
    unsigned long secondBytesRead = 0u;
    unsigned long restartBytesRead = 0u;
    unsigned long restartEofBytes  = 0u;
    std::vector<uint8_t> firstBytes;
    std::vector<uint8_t> restartBytes;
    EndpointSnapshot source;
};

struct CollectionScenarioResult final
{
    HRESULT setupHr     = E_FAIL;
    HRESULT operationHr = E_FAIL;
    uint64_t elapsedMs  = 0u;
    EndpointSnapshot source;
    EndpointSnapshot destination;
};

struct RenameScenarioResult final
{
    HRESULT setupHr     = E_FAIL;
    HRESULT operationHr = E_FAIL;
    uint64_t elapsedMs  = 0u;
    EndpointSnapshot endpoint;
};

class CleanupDebtAlertHost final : public IHost, public IHostAlerts
{
public:
    CleanupDebtAlertHost()                                            = default;
    CleanupDebtAlertHost(const CleanupDebtAlertHost&)                 = delete;
    CleanupDebtAlertHost(CleanupDebtAlertHost&&)                      = delete;
    CleanupDebtAlertHost& operator=(const CleanupDebtAlertHost&)      = delete;
    CleanupDebtAlertHost& operator=(CleanupDebtAlertHost&&)           = delete;
    ~CleanupDebtAlertHost()                                           = default;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** result) noexcept override
    {
        if (! result)
        {
            return E_POINTER;
        }

        *result = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IHost))
        {
            *result = static_cast<IHost*>(this);
        }
        else if (riid == __uuidof(IHostAlerts))
        {
            *result = static_cast<IHostAlerts*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        return _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
    }

    HRESULT STDMETHODCALLTYPE ShowAlert(const HostAlertRequest* request, void* /*cookie*/) noexcept override
    {
        if (! request || request->sizeBytes < sizeof(HostAlertRequest) || ! request->message)
        {
            return E_INVALIDARG;
        }

        std::scoped_lock lock(_mutex);
        CopyBoundedText(request->title, _title);
        CopyBoundedText(request->message, _message);
        _scope    = request->scope;
        _modality = request->modality;
        _severity = request->severity;
        _closable = request->closable;
        _alertCount.fetch_add(1u, std::memory_order_release);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ClearAlert(HostAlertScope /*scope*/, void* /*cookie*/) noexcept override
    {
        _clearCount.fetch_add(1u, std::memory_order_release);
        return S_OK;
    }

    [[nodiscard]] unsigned int AlertCount() const noexcept
    {
        return _alertCount.load(std::memory_order_acquire);
    }

    [[nodiscard]] std::wstring Message() const
    {
        std::scoped_lock lock(_mutex);
        return _message.data();
    }

    [[nodiscard]] bool HasExpectedWarningShape() const noexcept
    {
        std::scoped_lock lock(_mutex);
        return _scope == HOST_ALERT_SCOPE_APPLICATION && _modality == HOST_ALERT_MODELESS && _severity == HOST_ALERT_WARNING && _closable != FALSE &&
               _message[0] != L'\0';
    }

private:
    template <size_t Size>
    static void CopyBoundedText(const wchar_t* source, std::array<wchar_t, Size>& destination) noexcept
    {
        destination.fill(L'\0');
        if (! source)
        {
            return;
        }

        size_t length = 0u;
        while (length + 1u < Size && source[length] != L'\0')
        {
            ++length;
        }
        std::copy_n(source, length, destination.begin());
    }

    std::atomic<ULONG> _refCount{1u};
    mutable std::mutex _mutex;
    std::array<wchar_t, 256u> _title{};
    std::array<wchar_t, 1024u> _message{};
    HostAlertScope _scope       = HOST_ALERT_SCOPE_APPLICATION;
    HostAlertModality _modality = HOST_ALERT_MODELESS;
    HostAlertSeverity _severity = HOST_ALERT_INFO;
    BOOL _closable              = FALSE;
    std::atomic<unsigned int> _alertCount{0u};
    std::atomic<unsigned int> _clearCount{0u};
};

class CleanupDebtOperationCallback final : public IFileSystemCallback
{
public:
    CleanupDebtOperationCallback()                                                  = default;
    CleanupDebtOperationCallback(const CleanupDebtOperationCallback&)               = delete;
    CleanupDebtOperationCallback(CleanupDebtOperationCallback&&)                    = delete;
    CleanupDebtOperationCallback& operator=(const CleanupDebtOperationCallback&)    = delete;
    CleanupDebtOperationCallback& operator=(CleanupDebtOperationCallback&&)         = delete;
    ~CleanupDebtOperationCallback()                                                 = default;

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
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemItemCompleted(FileSystemOperation /*operationType*/,
                                                      unsigned long itemIndex,
                                                      const wchar_t* /*sourcePath*/,
                                                      const wchar_t* /*destinationPath*/,
                                                      HRESULT status,
                                                      const FileSystemItemMutationResult* mutationResult,
                                                      FileSystemOptions* /*options*/,
                                                      void* /*cookie*/) noexcept override
    {
        std::scoped_lock lock(_mutex);
        if (_completedCount < _statuses.size())
        {
            _indices[_completedCount]  = itemIndex;
            _statuses[_completedCount] = status;
            _mutations[_completedCount] = FileSystemRouteContract::SnapshotItemMutationResult(mutationResult);
            _invalidMutation = _invalidMutation || (mutationResult != nullptr && ! _mutations[_completedCount].has_value());
        }
        ++_completedCount;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* cancel, void* /*cookie*/) noexcept override
    {
        if (! cancel)
        {
            return E_POINTER;
        }
        *cancel = FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemIssue(FileSystemOperation /*operationType*/,
                                              const wchar_t* /*sourcePath*/,
                                              const wchar_t* /*destinationPath*/,
                                              HRESULT /*status*/,
                                              FileSystemIssueAction* action,
                                              IFileSystemBoundObject** expectedDestination,
                                              FileSystemOptions* /*options*/,
                                              void* /*cookie*/) noexcept override
    {
        if (! action || ! expectedDestination)
        {
            return E_POINTER;
        }
        *expectedDestination = nullptr;
        *action = FileSystemIssueAction::Cancel;
        return S_OK;
    }

    [[nodiscard]] bool CompletedSuccessfully(size_t expectedCount) const noexcept
    {
        std::scoped_lock lock(_mutex);
        if (_completedCount != expectedCount || expectedCount > _statuses.size())
        {
            return false;
        }

        for (size_t index = 0u; index < expectedCount; ++index)
        {
            if (_indices[index] >= expectedCount || _statuses[index] != S_OK)
            {
                return false;
            }
        }
        return true;
    }

    [[nodiscard]] bool SingleCompletionHasTruth(
        bool responseLost,
        FileSystemRouteContract::MutationClassification failureClass = FileSystemRouteContract::MutationClassification::Indeterminate) const noexcept
    {
        std::scoped_lock lock(_mutex);
        if (_completedCount != 1u || _indices[0] != 0u || _invalidMutation)
        {
            return false;
        }
        const auto* receipt = _mutations[0].has_value() ? &_mutations[0].value() : nullptr;
        if (responseLost)
        {
            return FAILED(_statuses[0]) && FileSystemRouteContract::ClassifyFailedMutation(true, _statuses[0], receipt) == failureClass;
        }
        return _statuses[0] == S_OK && receipt != nullptr && receipt->outcomeKnown == TRUE && receipt->mutationCommitted == TRUE &&
               receipt->originalStillPresent == FALSE;
    }

    [[nodiscard]] bool SingleCopyCompletionHasStatus(HRESULT expected) const noexcept
    {
        std::scoped_lock lock(_mutex);
        // Copy reports status, not a receipt asserting destructive source mutation.
        return _completedCount == 1u && _indices[0] == 0u && _statuses[0] == expected && ! _invalidMutation && ! _mutations[0].has_value();
    }

    [[nodiscard]] bool TwoCompletionsKeepIndependentTruth(HRESULT failureStatus = E_ACCESSDENIED) const noexcept
    {
        std::scoped_lock lock(_mutex);
        if (_completedCount != 2u || _invalidMutation || _indices[0] == _indices[1])
        {
            return false;
        }
        for (size_t index = 0u; index < 2u; ++index)
        {
            const auto* receipt = _mutations[index].has_value() ? &_mutations[index].value() : nullptr;
            if (_indices[index] == 0u)
            {
                if (_statuses[index] != S_OK || receipt == nullptr || receipt->outcomeKnown != TRUE || receipt->mutationCommitted != TRUE ||
                    receipt->originalStillPresent != FALSE)
                {
                    return false;
                }
            }
            else if (_indices[index] != 1u || _statuses[index] != failureStatus ||
                     FileSystemRouteContract::ClassifyFailedMutation(true, _statuses[index], receipt) !=
                         FileSystemRouteContract::MutationClassification::RetryableNoCommit)
            {
                return false;
            }
        }
        return true;
    }

private:
    mutable std::mutex _mutex;
    std::array<unsigned long, 8u> _indices{};
    std::array<HRESULT, 8u> _statuses{};
    std::array<std::optional<FileSystemItemMutationResult>, 8u> _mutations{};
    bool _invalidMutation  = false;
    size_t _completedCount = 0u;
};

class CleanupDebtWatchCallback final : public IFileSystemDirectoryWatchCallback
{
public:
    CleanupDebtWatchCallback()                                              = default;
    CleanupDebtWatchCallback(const CleanupDebtWatchCallback&)               = delete;
    CleanupDebtWatchCallback(CleanupDebtWatchCallback&&)                    = delete;
    CleanupDebtWatchCallback& operator=(const CleanupDebtWatchCallback&)    = delete;
    CleanupDebtWatchCallback& operator=(CleanupDebtWatchCallback&&)         = delete;
    ~CleanupDebtWatchCallback()                                             = default;

    HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const FileSystemDirectoryChangeNotification* notification,
                                                         void* /*cookie*/) noexcept override
    {
        if (! notification || notification->sizeBytes < sizeof(FileSystemDirectoryChangeNotification))
        {
            return E_INVALIDARG;
        }

        _notificationCount.fetch_add(1u, std::memory_order_acq_rel);
        _changeCount.fetch_add(notification->changeCount, std::memory_order_acq_rel);
        return S_OK;
    }

    [[nodiscard]] unsigned int NotificationCount() const noexcept
    {
        return _notificationCount.load(std::memory_order_acquire);
    }

private:
    std::atomic<unsigned int> _notificationCount{0u};
    std::atomic<unsigned int> _changeCount{0u};
};

struct CleanupDebtScenarioResult final
{
    HRESULT setupHr         = E_FAIL;
    HRESULT operationHr     = E_FAIL;
    HRESULT secondCommitHr  = E_FAIL;
    uint64_t elapsedMs      = 0u;
    size_t commandsAfterFirstCommit = 0u;
    unsigned int alertsAfterFirstCommit = 0u;
    unsigned int notificationsAfterFirstCommit = 0u;
    unsigned int alertCount        = 0u;
    unsigned int notificationCount = 0u;
    bool alertShapeValid            = false;
    bool callbacksSucceeded         = false;
    std::wstring alertMessage;
    EndpointSnapshot source;
    EndpointSnapshot destination;
};

[[nodiscard]] HRESULT CreateConfiguredFtpSelfTestFileSystem(wil::com_ptr<IFileSystem>& fileSystemOut, IHost* host = nullptr)
{
    fileSystemOut.reset();
    HRESULT hr = CreateFtpSelfTestInstance(fileSystemOut.put(), host);
    if (FAILED(hr) || ! fileSystemOut)
    {
        return FAILED(hr) ? hr : E_FAIL;
    }

    wil::com_ptr<IInformations> information;
    hr = fileSystemOut->QueryInterface(__uuidof(IInformations), information.put_void());
    if (FAILED(hr) || ! information)
    {
        fileSystemOut.reset();
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }

    constexpr char kConfiguration[] =
        R"({"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":1,"deleteMaxConcurrency":1,"ftpUseEpsv":true})";
    hr = information->SetConfiguration(kConfiguration);
    if (FAILED(hr))
    {
        fileSystemOut.reset();
    }
    return hr;
}

[[nodiscard]] std::vector<uint8_t> MakeSourceBytes()
{
    std::vector<uint8_t> bytes(4099u);
    for (size_t index = 0u; index < bytes.size(); ++index)
    {
        bytes[index] = static_cast<uint8_t>((index * 37u + 11u) & 0xffu);
    }
    return bytes;
}

[[nodiscard]] std::vector<uint8_t> MakeDestinationSentinel()
{
    constexpr std::string_view text = "existing-destination-must-survive-short-upload";
    return std::vector<uint8_t>(text.begin(), text.end());
}

[[nodiscard]] ScenarioResult RunScenario(UploadRetention retention,
                                         const std::vector<uint8_t>& sourceBytes,
                                         const std::vector<uint8_t>& destinationSentinel,
                                         bool move                    = true,
                                         bool destinationCanProveSize = true,
                                         DownloadDelivery sourceDelivery = DownloadDelivery::Complete,
                                         bool sourceCanProveSize          = true,
                                         std::optional<size_t> sourceListedSize = std::nullopt,
                                         bool sourceSizeAuthenticationFailure  = false,
                                         bool sourceSizeMissingFailure         = false)
{
    ScenarioResult result{};
    FakeFtpEndpoint source(UploadRetention::Complete, sourceCanProveSize, sourceCanProveSize, false, sourceDelivery);
    FakeFtpEndpoint destination(retention, destinationCanProveSize, destinationCanProveSize);
    source.SeedFile("/source.bin", sourceBytes);
    if (sourceListedSize.has_value())
    {
        source.SetListedSize("/source.bin", sourceListedSize.value());
    }
    source.SetSizeAuthenticationFailure(sourceSizeAuthenticationFailure);
    source.SetSizeMissingFailure(sourceSizeMissingFailure);
    destination.SeedFile("/destination.bin", destinationSentinel);

    HRESULT hr = source.Start();
    if (FAILED(hr))
    {
        result.setupHr = hr;
        return result;
    }
    hr = destination.Start();
    if (FAILED(hr))
    {
        result.setupHr = hr;
        return result;
    }

    wil::com_ptr<IFileSystem> fileSystem;
    hr = CreateFtpSelfTestInstance(fileSystem.put());
    if (FAILED(hr) || ! fileSystem)
    {
        result.setupHr = FAILED(hr) ? hr : E_FAIL;
        return result;
    }

    wil::com_ptr<IInformations> information;
    hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
    if (FAILED(hr) || ! information)
    {
        result.setupHr = FAILED(hr) ? hr : E_NOINTERFACE;
        return result;
    }

    constexpr char kConfiguration[] =
        R"({"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":1,"deleteMaxConcurrency":1,"ftpUseEpsv":true})";
    hr = information->SetConfiguration(kConfiguration);
    if (FAILED(hr))
    {
        result.setupHr = hr;
        return result;
    }

    const std::wstring sourcePath      = std::format(L"//anonymous@127.0.0.1:{}/source.bin", source.Port());
    const std::wstring destinationPath = std::format(L"//anonymous@127.0.0.1:{}/destination.bin", destination.Port());

    result.setupHr     = S_OK;
    const auto started = std::chrono::steady_clock::now();
    result.moveHr      = move ? fileSystem->MoveItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr)
                              : fileSystem->CopyItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr);
    result.elapsedMs   = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started).count());

    fileSystem.reset();
    information.reset();
    source.Stop();
    destination.Stop();
    result.source      = source.Snapshot();
    result.destination = destination.Snapshot();
    return result;
}

[[nodiscard]] CollectionScenarioResult RunBatchScenario(bool move,
                                                         const std::vector<uint8_t>& firstBytes,
                                                         const std::vector<uint8_t>& secondBytes,
                                                         DownloadDelivery delivery,
                                                         bool sourceCanProveSize,
                                                         bool staleListedSizes = false)
{
    CollectionScenarioResult result{};
    FakeFtpEndpoint source(UploadRetention::Complete, sourceCanProveSize, sourceCanProveSize, false, delivery);
    FakeFtpEndpoint destination(UploadRetention::Complete);
    source.SeedFile("/batch-a.bin", firstBytes);
    source.SeedFile("/batch-b.bin", secondBytes);
    if (staleListedSizes)
    {
        source.SetListedSize("/batch-a.bin", firstBytes.size() / 2u);
        source.SetListedSize("/batch-b.bin", secondBytes.size() / 2u);
    }
    destination.SeedDirectory("/batch-destination");

    HRESULT hr = source.Start();
    if (SUCCEEDED(hr))
    {
        hr = destination.Start();
    }

    wil::com_ptr<IFileSystem> fileSystem;
    if (SUCCEEDED(hr))
    {
        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
    }
    result.setupHr = hr;
    if (SUCCEEDED(hr))
    {
        const std::array<std::wstring, 2u> sourcePaths{
            std::format(L"//anonymous@127.0.0.1:{}/batch-a.bin", source.Port()),
            std::format(L"//anonymous@127.0.0.1:{}/batch-b.bin", source.Port()),
        };
        const std::array<const wchar_t*, 2u> sourcePathPointers{sourcePaths[0].c_str(), sourcePaths[1].c_str()};
        const std::wstring destinationFolder = std::format(L"//anonymous@127.0.0.1:{}/batch-destination", destination.Port());
        const FileSystemFlags flags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        result.operationHr = move ? fileSystem->MoveItems(sourcePathPointers.data(),
                                                          static_cast<unsigned long>(sourcePathPointers.size()),
                                                          destinationFolder.c_str(),
                                                          flags,
                                                          nullptr,
                                                          nullptr,
                                                          nullptr)
                                  : fileSystem->CopyItems(sourcePathPointers.data(),
                                                          static_cast<unsigned long>(sourcePathPointers.size()),
                                                          destinationFolder.c_str(),
                                                          flags,
                                                          nullptr,
                                                          nullptr,
                                                          nullptr);
    }

    fileSystem.reset();
    source.Stop();
    destination.Stop();
    result.source      = source.Snapshot();
    result.destination = destination.Snapshot();
    return result;
}

[[nodiscard]] CollectionScenarioResult RunRecursiveScenario(bool move,
                                                             const std::vector<uint8_t>& rootBytes,
                                                             const std::vector<uint8_t>& nestedBytes,
                                                             DownloadDelivery delivery,
                                                             bool sourceCanProveSize,
                                                             bool staleListedSizes = false)
{
    CollectionScenarioResult result{};
    FakeFtpEndpoint source(UploadRetention::Complete, sourceCanProveSize, sourceCanProveSize, false, delivery);
    FakeFtpEndpoint destination(UploadRetention::Complete);
    source.SeedDirectory("/source-tree/nested");
    source.SeedFile("/source-tree/root.bin", rootBytes);
    source.SeedFile("/source-tree/nested/child.bin", nestedBytes);
    if (staleListedSizes)
    {
        source.SetListedSize("/source-tree/root.bin", rootBytes.size() / 2u);
        source.SetListedSize("/source-tree/nested/child.bin", nestedBytes.size() / 2u);
    }

    HRESULT hr = source.Start();
    if (SUCCEEDED(hr))
    {
        hr = destination.Start();
    }

    wil::com_ptr<IFileSystem> fileSystem;
    if (SUCCEEDED(hr))
    {
        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
    }
    result.setupHr = hr;
    if (SUCCEEDED(hr))
    {
        const std::wstring sourcePath      = std::format(L"//anonymous@127.0.0.1:{}/source-tree", source.Port());
        const std::wstring destinationPath = std::format(L"//anonymous@127.0.0.1:{}/destination-tree", destination.Port());
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        result.operationHr = move ? fileSystem->MoveItem(sourcePath.c_str(), destinationPath.c_str(), flags, nullptr, nullptr, nullptr)
                                  : fileSystem->CopyItem(sourcePath.c_str(), destinationPath.c_str(), flags, nullptr, nullptr, nullptr);
    }

    fileSystem.reset();
    source.Stop();
    destination.Stop();
    result.source      = source.Snapshot();
    result.destination = destination.Snapshot();
    return result;
}

[[nodiscard]] ScenarioResult RunNoOverwriteSingleRaceScenario(bool move,
                                                               const std::vector<uint8_t>& sourceBytes,
                                                               const std::vector<uint8_t>& concurrentBytes)
{
    ScenarioResult result{};
    FakeFtpEndpoint source(UploadRetention::Complete);
    FakeFtpEndpoint destination(UploadRetention::Complete);
    source.SeedFile("/source.bin", sourceBytes);
    destination.SetFileBeforeRename("/destination.bin", concurrentBytes);

    HRESULT hr = source.Start();
    if (SUCCEEDED(hr))
    {
        hr = destination.Start();
    }

    wil::com_ptr<IFileSystem> fileSystem;
    if (SUCCEEDED(hr))
    {
        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
    }
    result.setupHr = hr;
    if (SUCCEEDED(hr))
    {
        const std::wstring sourcePath      = std::format(L"//anonymous@127.0.0.1:{}/source.bin", source.Port());
        const std::wstring destinationPath = std::format(L"//anonymous@127.0.0.1:{}/destination.bin", destination.Port());
        const auto started = std::chrono::steady_clock::now();
        result.moveHr = move ? fileSystem->MoveItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)
                             : fileSystem->CopyItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        result.elapsedMs = static_cast<uint64_t>(
            std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started).count());
    }

    fileSystem.reset();
    source.Stop();
    destination.Stop();
    result.source      = source.Snapshot();
    result.destination = destination.Snapshot();
    return result;
}

[[nodiscard]] CollectionScenarioResult RunNoOverwriteBatchRaceScenario(bool move,
                                                                        const std::vector<uint8_t>& firstBytes,
                                                                        const std::vector<uint8_t>& secondBytes,
                                                                        const std::vector<uint8_t>& concurrentBytes)
{
    CollectionScenarioResult result{};
    FakeFtpEndpoint source(UploadRetention::Complete);
    FakeFtpEndpoint destination(UploadRetention::Complete);
    source.SeedFile("/batch-a.bin", firstBytes);
    source.SeedFile("/batch-b.bin", secondBytes);
    destination.SeedDirectory("/batch-destination");
    destination.SetFileBeforeRename("/batch-destination/batch-a.bin", concurrentBytes);
    destination.SetFileBeforeRename("/batch-destination/batch-b.bin", concurrentBytes);

    HRESULT hr = source.Start();
    if (SUCCEEDED(hr))
    {
        hr = destination.Start();
    }

    wil::com_ptr<IFileSystem> fileSystem;
    if (SUCCEEDED(hr))
    {
        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
    }
    result.setupHr = hr;
    if (SUCCEEDED(hr))
    {
        const std::array<std::wstring, 2u> sourcePaths{
            std::format(L"//anonymous@127.0.0.1:{}/batch-a.bin", source.Port()),
            std::format(L"//anonymous@127.0.0.1:{}/batch-b.bin", source.Port()),
        };
        const std::array<const wchar_t*, 2u> sourcePathPointers{sourcePaths[0].c_str(), sourcePaths[1].c_str()};
        const std::wstring destinationFolder = std::format(L"//anonymous@127.0.0.1:{}/batch-destination", destination.Port());
        const auto started = std::chrono::steady_clock::now();
        result.operationHr = move ? fileSystem->MoveItems(sourcePathPointers.data(),
                                                          static_cast<unsigned long>(sourcePathPointers.size()),
                                                          destinationFolder.c_str(),
                                                          FILESYSTEM_FLAG_CONTINUE_ON_ERROR,
                                                          nullptr,
                                                          nullptr,
                                                          nullptr)
                                  : fileSystem->CopyItems(sourcePathPointers.data(),
                                                          static_cast<unsigned long>(sourcePathPointers.size()),
                                                          destinationFolder.c_str(),
                                                          FILESYSTEM_FLAG_CONTINUE_ON_ERROR,
                                                          nullptr,
                                                          nullptr,
                                                          nullptr);
        result.elapsedMs = static_cast<uint64_t>(
            std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started).count());
    }

    fileSystem.reset();
    source.Stop();
    destination.Stop();
    result.source      = source.Snapshot();
    result.destination = destination.Snapshot();
    return result;
}

[[nodiscard]] RenameScenarioResult RunNoOverwriteRenameRaceScenario(bool batch,
                                                                    const std::vector<uint8_t>& sourceBytes,
                                                                    const std::vector<uint8_t>& concurrentBytes)
{
    RenameScenarioResult result{};
    FakeFtpEndpoint endpoint(UploadRetention::Complete);
    endpoint.SeedFile("/rename-a.bin", sourceBytes);
    endpoint.SetFileBeforeRename("/renamed-a.bin", concurrentBytes);
    if (batch)
    {
        endpoint.SeedFile("/rename-b.bin", sourceBytes);
        endpoint.SetFileBeforeRename("/renamed-b.bin", concurrentBytes);
    }

    HRESULT hr = endpoint.Start();
    wil::com_ptr<IFileSystem> fileSystem;
    if (SUCCEEDED(hr))
    {
        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
    }
    result.setupHr = hr;
    if (SUCCEEDED(hr))
    {
        const std::wstring sourceA      = std::format(L"//anonymous@127.0.0.1:{}/rename-a.bin", endpoint.Port());
        const std::wstring destinationA = std::format(L"//anonymous@127.0.0.1:{}/renamed-a.bin", endpoint.Port());
        const auto started = std::chrono::steady_clock::now();
        if (! batch)
        {
            result.operationHr = fileSystem->RenameItem(sourceA.c_str(), destinationA.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        }
        else
        {
            const std::wstring sourceB = std::format(L"//anonymous@127.0.0.1:{}/rename-b.bin", endpoint.Port());
            const std::array<FileSystemRenamePair, 2u> pairs{
                FileSystemRenamePair{.sizeBytes = sizeof(FileSystemRenamePair), .sourcePath = sourceA.c_str(), .newName = L"renamed-a.bin"},
                FileSystemRenamePair{.sizeBytes = sizeof(FileSystemRenamePair), .sourcePath = sourceB.c_str(), .newName = L"renamed-b.bin"},
            };
            result.operationHr = fileSystem->RenameItems(pairs.data(),
                                                         static_cast<unsigned long>(pairs.size()),
                                                         FILESYSTEM_FLAG_CONTINUE_ON_ERROR,
                                                         nullptr,
                                                         nullptr,
                                                         nullptr);
        }
        result.elapsedMs = static_cast<uint64_t>(
            std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started).count());
    }

    fileSystem.reset();
    endpoint.Stop();
    result.endpoint = endpoint.Snapshot();
    return result;
}

[[nodiscard]] ScenarioResult RunMoveDeleteFailureRaceScenario(const std::vector<uint8_t>& sourceBytes,
                                                               const std::vector<uint8_t>& destinationSentinel,
                                                               const std::vector<uint8_t>& concurrentBytes)
{
    ScenarioResult result{};
    FakeFtpEndpoint source(UploadRetention::Complete);
    FakeFtpEndpoint destination(UploadRetention::Complete);
    source.SeedFile("/source.bin", sourceBytes);
    source.SetDeleteFailure("/source.bin");
    destination.SeedFile("/destination.bin", destinationSentinel);
    destination.SetFileAfterRename("/destination.bin", concurrentBytes);

    HRESULT hr = source.Start();
    if (SUCCEEDED(hr))
    {
        hr = destination.Start();
    }
    wil::com_ptr<IFileSystem> fileSystem;
    if (SUCCEEDED(hr))
    {
        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
    }
    result.setupHr = hr;
    if (SUCCEEDED(hr))
    {
        const std::wstring sourcePath      = std::format(L"//anonymous@127.0.0.1:{}/source.bin", source.Port());
        const std::wstring destinationPath = std::format(L"//anonymous@127.0.0.1:{}/destination.bin", destination.Port());
        const auto started = std::chrono::steady_clock::now();
        result.moveHr = fileSystem->MoveItem(
            sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr);
        result.elapsedMs = static_cast<uint64_t>(
            std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started).count());
    }

    fileSystem.reset();
    source.Stop();
    destination.Stop();
    result.source      = source.Snapshot();
    result.destination = destination.Snapshot();
    return result;
}

[[nodiscard]] CollectionScenarioResult RunRecursiveRollbackRaceScenario(const std::vector<uint8_t>& sourceBytes,
                                                                         const std::vector<uint8_t>& concurrentBytes)
{
    CollectionScenarioResult result{};
    FakeFtpEndpoint source(UploadRetention::Complete);
    FakeFtpEndpoint destination(UploadRetention::StrictPrefix);
    source.SeedDirectory("/source-tree");
    source.SeedFile("/source-tree/source.bin", sourceBytes);
    destination.SetFileAfterDirectoryCreate("/destination-tree", "/destination-tree/concurrent.bin", concurrentBytes);

    HRESULT hr = source.Start();
    if (SUCCEEDED(hr))
    {
        hr = destination.Start();
    }
    wil::com_ptr<IFileSystem> fileSystem;
    if (SUCCEEDED(hr))
    {
        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
    }
    result.setupHr = hr;
    if (SUCCEEDED(hr))
    {
        const std::wstring sourcePath      = std::format(L"//anonymous@127.0.0.1:{}/source-tree", source.Port());
        const std::wstring destinationPath = std::format(L"//anonymous@127.0.0.1:{}/destination-tree", destination.Port());
        const auto started = std::chrono::steady_clock::now();
        result.operationHr = fileSystem->CopyItem(sourcePath.c_str(),
                                                  destinationPath.c_str(),
                                                  static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                  nullptr,
                                                  nullptr,
                                                  nullptr);
        result.elapsedMs = static_cast<uint64_t>(
            std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started).count());
    }

    fileSystem.reset();
    source.Stop();
    destination.Stop();
    result.source      = source.Snapshot();
    result.destination = destination.Snapshot();
    return result;
}

[[nodiscard]] ReaderScenarioResult RunReaderScenario(const std::vector<uint8_t>& sourceBytes,
                                                      DownloadDelivery delivery,
                                                      std::optional<uint64_t> seekOffset = std::nullopt,
                                                      bool restartAfterFailure           = false,
                                                      std::optional<size_t> listedSize    = std::nullopt,
                                                      bool sizeAuthenticationFailure     = false,
                                                      bool sizeMissingFailure            = false)
{
    ReaderScenarioResult result{};
    FakeFtpEndpoint source(UploadRetention::Complete, true, true, false, delivery);
    source.SeedFile("/source.bin", sourceBytes);
    if (listedSize.has_value())
    {
        source.SetListedSize("/source.bin", listedSize.value());
    }
    source.SetSizeAuthenticationFailure(sizeAuthenticationFailure);
    source.SetSizeMissingFailure(sizeMissingFailure);

    HRESULT hr = source.Start();
    if (FAILED(hr))
    {
        result.setupHr = hr;
        return result;
    }

    wil::com_ptr<IFileSystem> fileSystem;
    hr = CreateFtpSelfTestInstance(fileSystem.put());
    if (FAILED(hr) || ! fileSystem)
    {
        result.setupHr = FAILED(hr) ? hr : E_FAIL;
        return result;
    }

    wil::com_ptr<IInformations> information;
    wil::com_ptr<IFileSystemIO> io;
    hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
    if (SUCCEEDED(hr))
    {
        hr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), io.put_void());
    }
    constexpr char kConfiguration[] =
        R"({"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":1,"deleteMaxConcurrency":1,"ftpUseEpsv":true})";
    if (SUCCEEDED(hr) && information)
    {
        hr = information->SetConfiguration(kConfiguration);
    }
    result.setupHr = hr;
    if (FAILED(hr) || ! io)
    {
        return result;
    }

    const std::wstring sourcePath = std::format(L"//anonymous@127.0.0.1:{}/source.bin", source.Port());
    wil::com_ptr<IFileReader> reader;
    result.createHr = io->CreateFileReader(sourcePath.c_str(), reader.put());
    if (FAILED(result.createHr) || ! reader)
    {
        source.Stop();
        result.source = source.Snapshot();
        return result;
    }

    result.getSizeHr = reader->GetSize(&result.committedSizeBytes);
    if (seekOffset.has_value())
    {
        result.seekHr = reader->Seek(static_cast<__int64>(seekOffset.value()), FILE_BEGIN, &result.seekPosition);
    }

    std::vector<uint8_t> buffer(sourceBytes.size() + 32u);
    result.firstReadHr = FAILED(result.seekHr) ? result.seekHr
                                               : reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &result.firstBytesRead);
    result.firstBytes.assign(buffer.begin(), buffer.begin() + static_cast<std::ptrdiff_t>(result.firstBytesRead));
    result.secondReadHr = FAILED(result.firstReadHr)
                              ? result.firstReadHr
                              : reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &result.secondBytesRead);

    if (restartAfterFailure)
    {
        source.SetDownloadDelivery(DownloadDelivery::Complete);
        uint64_t restartPosition = 99u;
        result.restartSeekHr     = reader->Seek(0, FILE_BEGIN, &restartPosition);
        result.restartReadHr = FAILED(result.restartSeekHr)
                                   ? result.restartSeekHr
                                   : reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &result.restartBytesRead);
        result.restartBytes.assign(buffer.begin(), buffer.begin() + static_cast<std::ptrdiff_t>(result.restartBytesRead));
        result.restartEofHr = FAILED(result.restartReadHr)
                                  ? result.restartReadHr
                                  : reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &result.restartEofBytes);
    }

    reader.reset();
    io.reset();
    information.reset();
    fileSystem.reset();
    source.Stop();
    result.source = source.Snapshot();
    return result;
}

[[nodiscard]] WriterLifetimeResult RunWriterScenario(const std::vector<uint8_t>& bytes,
                                                     UploadRetention retention,
                                                     const std::vector<uint8_t>* destinationSentinel,
                                                     FileSystemFlags flags,
                                                     bool commitWriter,
                                                     bool failWriterPromotion = false,
                                                     const std::vector<uint8_t>* concurrentDestinationBeforePromotion = nullptr)
{
    WriterLifetimeResult result{};
    FakeFtpEndpoint destination(retention, true, true, failWriterPromotion);
    if (destinationSentinel != nullptr)
    {
        destination.SeedFile("/owned-writer.bin", *destinationSentinel);
    }
    if (concurrentDestinationBeforePromotion != nullptr)
    {
        destination.SetFileBeforeRename("/owned-writer.bin", *concurrentDestinationBeforePromotion);
    }
    HRESULT hr = destination.Start();
    if (FAILED(hr))
    {
        result.setupHr = hr;
        return result;
    }

    wil::com_ptr<IFileSystem> fileSystem;
    hr = CreateFtpSelfTestInstance(fileSystem.put());
    if (FAILED(hr) || ! fileSystem)
    {
        result.setupHr = FAILED(hr) ? hr : E_FAIL;
        return result;
    }
    wil::com_ptr<IInformations> information;
    wil::com_ptr<IFileSystemIO> io;
    hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
    if (SUCCEEDED(hr))
    {
        hr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), io.put_void());
    }
    constexpr char kConfiguration[] =
        R"({"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":1,"deleteMaxConcurrency":1,"ftpUseEpsv":true})";
    if (SUCCEEDED(hr) && information)
    {
        hr = information->SetConfiguration(kConfiguration);
    }
    wil::com_ptr<IFileWriter> writer;
    const std::wstring destinationPath = std::format(L"//anonymous@127.0.0.1:{}/owned-writer.bin", destination.Port());
    if (SUCCEEDED(hr) && io)
    {
        result.createHr = io->CreateFileWriter(destinationPath.c_str(), flags, writer.put());
    }
    result.setupHr = hr;
    if (FAILED(hr) || FAILED(result.createHr) || ! writer)
    {
        destination.Stop();
        result.destination = destination.Snapshot();
        return result;
    }

    // The returned writer is now the only object retaining the filesystem instance.
    information.reset();
    io.reset();
    fileSystem.reset();
    const auto started = std::chrono::steady_clock::now();
    result.writeHr     = writer->Write(bytes.data(), static_cast<unsigned long>(bytes.size()), &result.written);
    if (commitWriter)
    {
        result.commitHr = SUCCEEDED(result.writeHr) ? writer->Commit() : result.writeHr;
    }
    else
    {
        result.commitHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    writer.reset();
    result.elapsedMs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started).count());
    destination.Stop();
    result.destination = destination.Snapshot();
    return result;
}

[[nodiscard]] CleanupDebtScenarioResult RunCleanupDebtWriterScenario(const std::vector<uint8_t>& bytes,
                                                                     const std::vector<uint8_t>& destinationSentinel,
                                                                     UploadRetention retention = UploadRetention::Complete,
                                                                     bool rejectWriterStagingCleanup = false,
                                                                     bool retryCommit = true)
{
    CleanupDebtScenarioResult result{};
    CleanupDebtAlertHost host;
    CleanupDebtWatchCallback watchCallback;
    FakeFtpEndpoint destination(retention);
    destination.SeedFile("/owned-writer.bin", destinationSentinel);
    destination.SetDeleteFailureForMarker(std::string(kRollbackSiblingMarker));
    if (rejectWriterStagingCleanup)
    {
        destination.SetDeleteFailureForMarker(std::string(kWriterSiblingMarker));
    }

    HRESULT hr = destination.Start();
    wil::com_ptr<IFileSystem> fileSystem;
    if (SUCCEEDED(hr))
    {
        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem, &host);
    }

    wil::com_ptr<IFileSystemIO> io;
    wil::com_ptr<IFileSystemDirectoryWatch> watch;
    if (SUCCEEDED(hr))
    {
        hr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), io.put_void());
    }
    if (SUCCEEDED(hr))
    {
        hr = fileSystem->QueryInterface(__uuidof(IFileSystemDirectoryWatch), watch.put_void());
    }

    const std::wstring watchPath       = std::format(L"//anonymous@127.0.0.1:{}/", destination.Port());
    const std::wstring destinationPath = std::format(L"{}owned-writer.bin", watchPath);
    if (SUCCEEDED(hr))
    {
        hr = watch->WatchDirectory(watchPath.c_str(), &watchCallback, nullptr);
    }

    wil::com_ptr<IFileWriter> writer;
    if (SUCCEEDED(hr))
    {
        result.operationHr = io->CreateFileWriter(destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.put());
        if (FAILED(result.operationHr) || ! writer)
        {
            hr = FAILED(result.operationHr) ? result.operationHr : E_FAIL;
        }
    }

    unsigned long written = 0u;
    const auto started     = std::chrono::steady_clock::now();
    if (SUCCEEDED(hr))
    {
        hr = writer->Write(bytes.data(), static_cast<unsigned long>(bytes.size()), &written);
    }
    if (SUCCEEDED(hr))
    {
        result.operationHr = writer->Commit();
        result.commandsAfterFirstCommit      = destination.Snapshot().commands.size();
        result.alertsAfterFirstCommit        = host.AlertCount();
        result.notificationsAfterFirstCommit = watchCallback.NotificationCount();
        result.secondCommitHr                = retryCommit ? writer->Commit() : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    result.elapsedMs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started).count());

    if (watch)
    {
        static_cast<void>(watch->UnwatchDirectory(watchPath.c_str()));
    }
    writer.reset();
    watch.reset();
    io.reset();
    fileSystem.reset();
    destination.Stop();

    result.setupHr           = hr;
    result.alertCount        = host.AlertCount();
    result.notificationCount = watchCallback.NotificationCount();
    result.alertShapeValid   = host.HasExpectedWarningShape();
    result.alertMessage      = host.Message();
    result.callbacksSucceeded = written == bytes.size();
    result.destination       = destination.Snapshot();
    return result;
}

enum class CleanupDebtTransferShape
{
    Single,
    Recursive,
    Batch,
};

[[nodiscard]] CleanupDebtScenarioResult RunCleanupDebtTransferScenario(CleanupDebtTransferShape shape,
                                                                       bool move,
                                                                       const std::vector<uint8_t>& firstBytes,
                                                                       const std::vector<uint8_t>& secondBytes,
                                                                       const std::vector<uint8_t>& destinationSentinel)
{
    CleanupDebtScenarioResult result{};
    CleanupDebtAlertHost host;
    CleanupDebtOperationCallback operationCallback;
    CleanupDebtWatchCallback watchCallback;
    FakeFtpEndpoint source(UploadRetention::Complete);
    FakeFtpEndpoint destination(UploadRetention::Complete);

    switch (shape)
    {
        case CleanupDebtTransferShape::Single:
            source.SeedFile("/source.bin", firstBytes);
            destination.SeedFile("/destination.bin", destinationSentinel);
            break;
        case CleanupDebtTransferShape::Recursive:
            source.SeedDirectory("/source-tree/nested");
            source.SeedFile("/source-tree/root.bin", firstBytes);
            source.SeedFile("/source-tree/nested/child.bin", secondBytes);
            destination.SeedDirectory("/destination-tree/nested");
            destination.SeedFile("/destination-tree/root.bin", destinationSentinel);
            destination.SeedFile("/destination-tree/nested/child.bin", destinationSentinel);
            break;
        case CleanupDebtTransferShape::Batch:
            source.SeedFile("/batch-a.bin", firstBytes);
            source.SeedFile("/batch-b.bin", secondBytes);
            destination.SeedDirectory("/batch-destination");
            destination.SeedFile("/batch-destination/batch-a.bin", destinationSentinel);
            destination.SeedFile("/batch-destination/batch-b.bin", destinationSentinel);
            break;
    }
    destination.SetDeleteFailureForMarker(std::string(kRollbackSiblingMarker));

    HRESULT hr = source.Start();
    if (SUCCEEDED(hr))
    {
        hr = destination.Start();
    }

    wil::com_ptr<IFileSystem> fileSystem;
    if (SUCCEEDED(hr))
    {
        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem, &host);
    }

    wil::com_ptr<IFileSystemDirectoryWatch> watch;
    if (SUCCEEDED(hr))
    {
        hr = fileSystem->QueryInterface(__uuidof(IFileSystemDirectoryWatch), watch.put_void());
    }

    const std::wstring destinationRoot = std::format(L"//anonymous@127.0.0.1:{}", destination.Port());
    const std::wstring watchPath =
        shape == CleanupDebtTransferShape::Batch ? std::format(L"{}/batch-destination", destinationRoot) : std::format(L"{}/", destinationRoot);
    if (SUCCEEDED(hr))
    {
        hr = watch->WatchDirectory(watchPath.c_str(), &watchCallback, nullptr);
    }

    result.setupHr = hr;
    const auto started = std::chrono::steady_clock::now();
    if (SUCCEEDED(hr))
    {
        const FileSystemFlags overwriteFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                                            (shape == CleanupDebtTransferShape::Recursive
                                                                                 ? FILESYSTEM_FLAG_RECURSIVE
                                                                                 : FILESYSTEM_FLAG_NONE));
        if (shape == CleanupDebtTransferShape::Single)
        {
            const std::wstring sourcePath      = std::format(L"//anonymous@127.0.0.1:{}/source.bin", source.Port());
            const std::wstring destinationPath = std::format(L"{}/destination.bin", destinationRoot);
            result.operationHr = move ? fileSystem->MoveItem(sourcePath.c_str(),
                                                              destinationPath.c_str(),
                                                              overwriteFlags,
                                                              nullptr,
                                                              &operationCallback,
                                                              nullptr)
                                      : fileSystem->CopyItem(sourcePath.c_str(),
                                                              destinationPath.c_str(),
                                                              overwriteFlags,
                                                              nullptr,
                                                              &operationCallback,
                                                              nullptr);
        }
        else if (shape == CleanupDebtTransferShape::Recursive)
        {
            const std::wstring sourcePath      = std::format(L"//anonymous@127.0.0.1:{}/source-tree", source.Port());
            const std::wstring destinationPath = std::format(L"{}/destination-tree", destinationRoot);
            result.operationHr = move ? fileSystem->MoveItem(sourcePath.c_str(),
                                                              destinationPath.c_str(),
                                                              overwriteFlags,
                                                              nullptr,
                                                              &operationCallback,
                                                              nullptr)
                                      : fileSystem->CopyItem(sourcePath.c_str(),
                                                              destinationPath.c_str(),
                                                              overwriteFlags,
                                                              nullptr,
                                                              &operationCallback,
                                                              nullptr);
        }
        else
        {
            const std::array<std::wstring, 2u> sourcePaths{
                std::format(L"//anonymous@127.0.0.1:{}/batch-a.bin", source.Port()),
                std::format(L"//anonymous@127.0.0.1:{}/batch-b.bin", source.Port()),
            };
            const std::array<const wchar_t*, 2u> sourcePathPointers{sourcePaths[0].c_str(), sourcePaths[1].c_str()};
            const std::wstring destinationFolder = std::format(L"{}/batch-destination", destinationRoot);
            const FileSystemFlags batchFlags = static_cast<FileSystemFlags>(overwriteFlags | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
            result.operationHr = move ? fileSystem->MoveItems(sourcePathPointers.data(),
                                                               static_cast<unsigned long>(sourcePathPointers.size()),
                                                               destinationFolder.c_str(),
                                                               batchFlags,
                                                               nullptr,
                                                               &operationCallback,
                                                               nullptr)
                                      : fileSystem->CopyItems(sourcePathPointers.data(),
                                                               static_cast<unsigned long>(sourcePathPointers.size()),
                                                               destinationFolder.c_str(),
                                                               batchFlags,
                                                               nullptr,
                                                               &operationCallback,
                                                               nullptr);
        }
    }
    result.elapsedMs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started).count());

    if (watch)
    {
        static_cast<void>(watch->UnwatchDirectory(watchPath.c_str()));
    }
    watch.reset();
    fileSystem.reset();
    source.Stop();
    destination.Stop();

    const size_t expectedCallbacks = shape == CleanupDebtTransferShape::Batch ? 2u : 1u;
    result.alertCount        = host.AlertCount();
    result.notificationCount = watchCallback.NotificationCount();
    result.alertShapeValid   = host.HasExpectedWarningShape();
    result.alertMessage      = host.Message();
    result.callbacksSucceeded = operationCallback.CompletedSuccessfully(expectedCallbacks);
    result.source            = source.Snapshot();
    result.destination       = destination.Snapshot();
    return result;
}

[[nodiscard]] bool FileEquals(const EndpointSnapshot& snapshot, std::string_view path, const std::vector<uint8_t>& expected)
{
    const auto found = snapshot.files.find(std::string(path));
    return found != snapshot.files.end() && found->second == expected;
}

[[nodiscard]] bool FileAbsent(const EndpointSnapshot& snapshot, std::string_view path)
{
    return snapshot.files.find(std::string(path)) == snapshot.files.end();
}

[[nodiscard]] bool HasNoTransactionSiblings(const EndpointSnapshot& snapshot) noexcept
{
    for (const auto& [path, bytes] : snapshot.files)
    {
        static_cast<void>(bytes);
        if (path.find(kUploadSiblingMarker) != std::string::npos || path.find(kRollbackSiblingMarker) != std::string::npos ||
            path.find(kWriterSiblingMarker) != std::string::npos)
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] size_t CountCommands(const EndpointSnapshot& snapshot, std::string_view verb, std::string_view argumentFragment = {}) noexcept
{
    size_t count = 0u;
    for (const CommandRecord& command : snapshot.commands)
    {
        if (command.verb == verb && (argumentFragment.empty() || command.argument.find(argumentFragment) != std::string::npos))
        {
            ++count;
        }
    }
    return count;
}

[[nodiscard]] bool DeletedTransactionSibling(const EndpointSnapshot& snapshot, std::string_view marker) noexcept
{
    return std::ranges::any_of(snapshot.deletedPaths, [marker](const std::string& path) noexcept { return path.find(marker) != std::string::npos; });
}

[[nodiscard]] bool TransactionSiblingEquals(const EndpointSnapshot& snapshot,
                                            std::string_view marker,
                                            const std::vector<uint8_t>& expected) noexcept
{
    return std::ranges::any_of(snapshot.files, [marker, &expected](const auto& entry) noexcept {
        return entry.first.find(marker) != std::string::npos && entry.second == expected;
    });
}

[[nodiscard]] size_t CountTransactionSiblings(const EndpointSnapshot& snapshot, std::string_view marker) noexcept
{
    return static_cast<size_t>(std::ranges::count_if(snapshot.files, [marker](const auto& entry) noexcept {
        return entry.first.find(marker) != std::string::npos;
    }));
}

[[nodiscard]] bool CleanupDebtAlertIsContentFree(const CleanupDebtScenarioResult& result, size_t expectedCount)
{
    const std::wstring expectedCountText = std::to_wstring(expectedCount);
    constexpr std::array<std::wstring_view, 8u> kForbiddenFragments{{
        L"127.0.0.1", L"anonymous", L".bin", L"source-tree", L"destination-tree", L".redsalamander-", L"//", L"\\",
    }};

    return ! result.alertMessage.empty() && result.alertMessage.find(expectedCountText) != std::wstring::npos &&
           std::ranges::none_of(kForbiddenFragments, [&result](std::wstring_view fragment) {
               return result.alertMessage.find(fragment) != std::wstring::npos;
           });
}

void CheckShortScenario(const ScenarioResult& result,
                        const std::vector<uint8_t>& sourceBytes,
                        const std::vector<uint8_t>& destinationSentinel,
                        unsigned int& passed,
                        unsigned int& failed) noexcept
{
    DebugCheck(result.setupHr == S_OK, L"short-success fixture setup should succeed", passed, failed);
    DebugCheck(result.moveHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), L"short successful upload should return ERROR_PARTIAL_COPY", passed, failed);
    DebugCheck(FileEquals(result.source, "/source.bin", sourceBytes), L"short successful MOVE should preserve the source bytes", passed, failed);
    DebugCheck(FileEquals(result.destination, "/destination.bin", destinationSentinel),
               L"short successful MOVE should leave the pre-existing destination unchanged",
               passed,
               failed);
    DebugCheck(HasNoTransactionSiblings(result.destination), L"short successful MOVE should remove every staged or rollback sibling", passed, failed);
    DebugCheck(result.destination.lastUploadReceivedBytes == sourceBytes.size() &&
                   result.destination.lastUploadRetainedBytes < result.destination.lastUploadReceivedBytes,
               L"fake destination should consume the full upload while retaining a strict prefix",
               passed,
               failed);
    DebugCheck(CountCommands(result.destination, "STOR", kUploadSiblingMarker) == 1u,
               L"short-success proof should execute exactly one real Curl staged STOR",
               passed,
               failed);
    DebugCheck(CountCommands(result.destination, "SIZE", kUploadSiblingMarker) == 1u,
               L"short-success proof should execute exactly one successful staged SIZE",
               passed,
               failed);
    DebugCheck(
        DeletedTransactionSibling(result.destination, kUploadSiblingMarker), L"short-success proof should delete the staged remote object", passed, failed);
    DebugCheck(CountCommands(result.source, "DELE", "source.bin") == 0u, L"short-success proof should not issue source deletion", passed, failed);
    DebugCheck(SUCCEEDED(result.source.serverHr) && SUCCEEDED(result.destination.serverHr),
               L"short-success fake endpoints should finish without transport-fixture failures",
               passed,
               failed);
    DebugCheck(result.elapsedMs <= 15000u && result.source.commands.size() <= 64u && result.destination.commands.size() <= 96u,
               L"short-success proof should remain bounded by elapsed time and command counts",
               passed,
               failed);
}

void CheckControlScenario(const ScenarioResult& result,
                          const std::vector<uint8_t>& sourceBytes,
                          const std::vector<uint8_t>& destinationSentinel,
                          unsigned int& passed,
                          unsigned int& failed) noexcept
{
    DebugCheck(result.setupHr == S_OK, L"complete-upload control fixture setup should succeed", passed, failed);
    DebugCheck(result.moveHr == S_OK, L"complete-upload control MOVE should succeed", passed, failed);
    DebugCheck(FileAbsent(result.source, "/source.bin"), L"complete-upload control MOVE should delete the source", passed, failed);
    DebugCheck(
        FileEquals(result.destination, "/destination.bin", sourceBytes), L"complete-upload control MOVE should promote the exact source bytes", passed, failed);
    DebugCheck(TransactionSiblingEquals(result.destination, kRollbackSiblingMarker, destinationSentinel),
               L"complete-upload control MOVE should preserve its rollback backup without identity-safe conditional delete",
               passed,
               failed);
    DebugCheck(result.destination.lastUploadReceivedBytes == sourceBytes.size() && result.destination.lastUploadRetainedBytes == sourceBytes.size(),
               L"complete-upload control should retain the complete Curl stream",
               passed,
               failed);
    DebugCheck(CountCommands(result.destination, "STOR", kUploadSiblingMarker) == 1u && CountCommands(result.destination, "SIZE", kUploadSiblingMarker) == 1u,
               L"complete-upload control should use one staged STOR and one staged SIZE",
               passed,
               failed);
    DebugCheck(
        CountCommands(result.source, "DELE", "source.bin") == 1u, L"complete-upload control should delete the source only after promotion", passed, failed);
    DebugCheck(! DeletedTransactionSibling(result.destination, kRollbackSiblingMarker),
               L"complete-upload overwrite control should never delete its rollback sibling by pathname",
               passed,
               failed);
    DebugCheck(SUCCEEDED(result.source.serverHr) && SUCCEEDED(result.destination.serverHr),
               L"complete-upload fake endpoints should finish without transport-fixture failures",
               passed,
               failed);
    DebugCheck(result.elapsedMs <= 15000u && result.source.commands.size() <= 64u && result.destination.commands.size() <= 128u,
               L"complete-upload control should remain bounded by elapsed time and command counts",
               passed,
               failed);
}

void CheckUnknownSizeScenario(const ScenarioResult& result,
                              const std::vector<uint8_t>& sourceBytes,
                              const std::vector<uint8_t>& destinationSentinel,
                              std::wstring_view operation,
                              unsigned int& passed,
                              unsigned int& failed) noexcept
{
    DebugCheck(result.setupHr == S_OK, L"unknown-size fixture setup should succeed", passed, failed);
    DebugCheck(result.moveHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
               std::format(L"unknown-size {} should fail closed with ERROR_PARTIAL_COPY", operation).c_str(),
               passed,
               failed);
    DebugCheck(
        FileEquals(result.source, "/source.bin", sourceBytes), std::format(L"unknown-size {} should preserve the source", operation).c_str(), passed, failed);
    DebugCheck(FileEquals(result.destination, "/destination.bin", destinationSentinel),
               std::format(L"unknown-size {} should leave the destination unchanged", operation).c_str(),
               passed,
               failed);
    DebugCheck(HasNoTransactionSiblings(result.destination) && DeletedTransactionSibling(result.destination, kUploadSiblingMarker),
               std::format(L"unknown-size {} should remove the staged sibling", operation).c_str(),
               passed,
               failed);
    DebugCheck(CountCommands(result.destination, "SIZE", kUploadSiblingMarker) >= 1u &&
                   CountCommands(result.destination, "LIST") + CountCommands(result.destination, "NLST") >= 1u,
               std::format(L"unknown-size {} should exhaust both remote size proof routes", operation).c_str(),
               passed,
               failed);
    DebugCheck(CountCommands(result.source, "DELE", "source.bin") == 0u,
               std::format(L"unknown-size {} should never delete the intact source", operation).c_str(),
               passed,
               failed);
}

// Deterministic loopback IMAP4rev1 server reached through the real libcurl transport.
// Unlike ImapListingFixture (a synchronous request hook), this proves which libcurl
// channel actually carries `* N FETCH` rows for listing-shaped and bare UID sets.
struct FakeImapMessage final
{
    uint64_t uid = 0u;
    std::string subjectUtf8;
    bool literalSubject = false;
    uint64_t sizeBytes  = 0u;
};

struct ImapEndpointSnapshot final
{
    std::vector<CommandRecord> commands;
    std::vector<std::string> fetchSets;
    size_t connections = 0u;
    size_t logins      = 0u;
    HRESULT serverHr   = S_OK;
};

class FakeImapEndpoint final
{
public:
    FakeImapEndpoint() = default;
    ~FakeImapEndpoint() noexcept
    {
        Stop();
    }

    FakeImapEndpoint(const FakeImapEndpoint&)            = delete;
    FakeImapEndpoint& operator=(const FakeImapEndpoint&) = delete;
    FakeImapEndpoint(FakeImapEndpoint&&)                 = delete;
    FakeImapEndpoint& operator=(FakeImapEndpoint&&)      = delete;

    void AddMessage(FakeImapMessage message)
    {
        std::scoped_lock lock(_stateMutex);
        _messages.push_back(std::move(message));
        std::sort(_messages.begin(), _messages.end(), [](const auto& left, const auto& right) noexcept { return left.uid < right.uid; });
    }

    // The row for this UID stops after an unfinished `{N}` literal, which is what a
    // server disconnect or a damaged response looks like to the parser.
    void SetDamagedFetchUid(uint64_t uid) noexcept
    {
        std::scoped_lock lock(_stateMutex);
        _damagedUid = uid;
    }

    [[nodiscard]] HRESULT Start()
    {
        HRESULT hr = CreateLoopbackEndpoint(_listener, _port);
        if (FAILED(hr))
        {
            return hr;
        }
        _thread = std::jthread([this](std::stop_token stopToken) noexcept
        {
            try
            {
                ServerMain(stopToken);
            }
            catch (const std::bad_alloc&)
            {
                std::terminate();
            }
            catch (const std::exception&)
            {
                // Mandatory thread boundary: convert an unexpected fixture exception into deterministic test failure.
                Debug::Error(L"FileSystemCurl fake IMAP endpoint terminated after std::exception.");
                SetServerFailure(E_FAIL);
            }
        });
        return S_OK;
    }

    void Stop() noexcept
    {
        if (_thread.joinable())
        {
            _thread.request_stop();
            _thread.join();
        }
        _listener.reset();
    }

    [[nodiscard]] unsigned short Port() const noexcept
    {
        return _port;
    }

    [[nodiscard]] ImapEndpointSnapshot Snapshot() const
    {
        std::scoped_lock lock(_stateMutex);
        ImapEndpointSnapshot snapshot;
        snapshot.commands    = _commands;
        snapshot.fetchSets   = _fetchSets;
        snapshot.connections = _connections;
        snapshot.logins      = _logins;
        snapshot.serverHr    = _serverHr;
        return snapshot;
    }

private:
    void SetServerFailure(HRESULT hr) noexcept
    {
        std::scoped_lock lock(_stateMutex);
        if (SUCCEEDED(_serverHr))
        {
            _serverHr = hr;
        }
    }

    void ServerMain(std::stop_token stopToken)
    {
        std::vector<std::jthread> sessions;
        const auto stopSessions = wil::scope_exit([&sessions]() noexcept
        {
            for (std::jthread& session : sessions)
            {
                session.request_stop();
            }
            sessions.clear(); // joins
        });
        while (! stopToken.stop_requested())
        {
            const HRESULT waitHr = WaitForReadable(_listener.get(), stopToken, std::chrono::steady_clock::now() + 24h);
            if (FAILED(waitHr))
            {
                if (waitHr != HRESULT_FROM_WIN32(ERROR_CANCELLED))
                {
                    SetServerFailure(waitHr);
                }
                return;
            }
            wil::unique_socket client(accept(_listener.get(), nullptr, nullptr));
            if (! client)
            {
                SetServerFailure(SocketErrorToHResult());
                return;
            }
            {
                std::scoped_lock lock(_stateMutex);
                ++_connections;
            }
            sessions.emplace_back([this, client = std::move(client)](std::stop_token sessionStop) mutable noexcept
            {
                try
                {
                    Session(client.get(), sessionStop);
                }
                catch (const std::bad_alloc&)
                {
                    std::terminate();
                }
                catch (const std::exception&)
                {
                    // Mandatory thread boundary: a fixture session failure becomes a test failure, not a crash.
                    Debug::Error(L"FileSystemCurl fake IMAP session terminated after std::exception.");
                    SetServerFailure(E_FAIL);
                }
            });
        }
    }

    [[nodiscard]] static std::string_view TrimSpaces(std::string_view value) noexcept
    {
        while (! value.empty() && (value.front() == ' ' || value.front() == '\t'))
        {
            value.remove_prefix(1u);
        }
        while (! value.empty() && (value.back() == ' ' || value.back() == '\t'))
        {
            value.remove_suffix(1u);
        }
        return value;
    }

    [[nodiscard]] std::string BuildFetchRows(std::string_view uidSet)
    {
        std::string rows;
        std::scoped_lock lock(_stateMutex);
        _fetchSets.emplace_back(uidSet);
        while (! uidSet.empty())
        {
            const size_t comma           = uidSet.find(',');
            const std::string_view token = uidSet.substr(0u, comma);
            const size_t colon           = token.find(':');
            uint64_t first               = 0u;
            uint64_t last                = 0u;
            const std::string_view firstText = token.substr(0u, colon);
            const std::string_view lastText  = colon == std::string_view::npos ? token : token.substr(colon + 1u);
            const auto firstParsed = std::from_chars(firstText.data(), firstText.data() + firstText.size(), first);
            if (firstParsed.ec != std::errc{} || firstParsed.ptr != firstText.data() + firstText.size())
            {
                return {};
            }
            if (lastText == "*")
            {
                last = _messages.empty() ? first : _messages.back().uid;
            }
            else
            {
                const auto lastParsed = std::from_chars(lastText.data(), lastText.data() + lastText.size(), last);
                if (lastParsed.ec != std::errc{} || lastParsed.ptr != lastText.data() + lastText.size())
                {
                    return {};
                }
            }
            for (size_t index = 0u; index < _messages.size(); ++index)
            {
                const FakeImapMessage& message = _messages[index];
                if (message.uid < first || message.uid > last)
                {
                    continue;
                }
                rows.append(std::format("* {} FETCH (UID {} FLAGS (\\Seen) INTERNALDATE \"05-Sep-2026 10:00:00 +0000\" RFC822.SIZE {} ENVELOPE (NIL ",
                                        index + 1u,
                                        message.uid,
                                        message.sizeBytes));
                if (_damagedUid.has_value() && _damagedUid.value() == message.uid)
                {
                    // Announce a 64-byte literal but deliver only its first bytes.
                    rows.append("{64}\r\ndamaged\r\n");
                    continue;
                }
                if (message.literalSubject)
                {
                    rows.append(std::format("{{{}}}\r\n{}", message.subjectUtf8.size(), message.subjectUtf8));
                }
                else
                {
                    rows.append(std::format("\"{}\"", message.subjectUtf8));
                }
                rows.append(" NIL NIL NIL NIL NIL NIL NIL NIL))\r\n");
            }
            if (comma == std::string_view::npos)
            {
                break;
            }
            uidSet.remove_prefix(comma + 1u);
        }
        return rows;
    }

    void Session(SOCKET client, std::stop_token stopToken)
    {
        HRESULT hr = ConfigureSocketTimeouts(client);
        if (FAILED(hr))
        {
            SetServerFailure(hr);
            return;
        }
        hr = SendAll(client, "* OK [CAPABILITY IMAP4rev1] RedSalamander fake IMAP ready\r\n");
        if (FAILED(hr))
        {
            SetServerFailure(hr);
            return;
        }
        std::string pending;
        std::string line;
        while (! stopToken.stop_requested())
        {
            hr = ReadControlLine(client, stopToken, pending, line);
            if (FAILED(hr))
            {
                // Idle close and cancellation are ordinary session ends for a pooled client.
                if (hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) && hr != HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED) &&
                    hr != HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT) && hr != HRESULT_FROM_WIN32(WSAECONNRESET))
                {
                    SetServerFailure(hr);
                }
                return;
            }
            const std::string_view request = TrimSpaces(line);
            const size_t tagEnd            = request.find(' ');
            if (tagEnd == std::string_view::npos)
            {
                hr = SendAll(client, "* BAD missing command\r\n");
                if (FAILED(hr))
                {
                    SetServerFailure(hr);
                    return;
                }
                continue;
            }
            const std::string tag(request.substr(0u, tagEnd));
            std::string_view rest = TrimSpaces(request.substr(tagEnd + 1u));
            const size_t verbEnd  = rest.find(' ');
            std::string verb      = UpperAscii(rest.substr(0u, verbEnd));
            std::string argument(verbEnd == std::string_view::npos ? std::string_view{} : TrimSpaces(rest.substr(verbEnd + 1u)));
            if (verb == "UID")
            {
                const size_t subEnd = argument.find(' ');
                verb.append(" ").append(UpperAscii(std::string_view(argument).substr(0u, subEnd)));
                argument = subEnd == std::string::npos ? std::string{} : std::string(TrimSpaces(std::string_view(argument).substr(subEnd + 1u)));
            }
            {
                std::scoped_lock lock(_stateMutex);
                _commands.push_back(CommandRecord{.verb = verb, .argument = argument});
            }

            std::string reply;
            bool logout = false;
            if (verb == "CAPABILITY")
            {
                reply = std::format("* CAPABILITY IMAP4rev1\r\n{} OK CAPABILITY completed\r\n", tag);
            }
            else if (verb == "LOGIN")
            {
                {
                    std::scoped_lock lock(_stateMutex);
                    ++_logins;
                }
                reply = std::format("{} OK [CAPABILITY IMAP4rev1] LOGIN completed\r\n", tag);
            }
            else if (verb == "LIST" || verb == "LSUB")
            {
                reply = std::format("* LIST () \"/\" \"INBOX\"\r\n{} OK LIST completed\r\n", tag);
            }
            else if (verb == "STATUS")
            {
                std::scoped_lock lock(_stateMutex);
                const uint64_t nextUid = _messages.empty() ? 1u : _messages.back().uid + 1u;
                reply                  = std::format("* STATUS \"INBOX\" (MESSAGES {} RECENT 0 UIDNEXT {} UIDVALIDITY 777 UNSEEN 0)\r\n{} OK STATUS completed\r\n",
                                    _messages.size(),
                                    nextUid,
                                    tag);
            }
            else if (verb == "SELECT" || verb == "EXAMINE")
            {
                std::scoped_lock lock(_stateMutex);
                const uint64_t nextUid = _messages.empty() ? 1u : _messages.back().uid + 1u;
                reply                  = std::format("* {} EXISTS\r\n* 0 RECENT\r\n* FLAGS (\\Seen \\Flagged \\Deleted)\r\n"
                                                     "* OK [UIDVALIDITY 777] UIDs valid\r\n* OK [UIDNEXT {}] Predicted next UID\r\n{} OK [READ-WRITE] {} completed\r\n",
                                    _messages.size(),
                                    nextUid,
                                    tag,
                                    verb);
            }
            else if (verb == "UID SEARCH")
            {
                std::scoped_lock lock(_stateMutex);
                reply = "* SEARCH";
                for (const FakeImapMessage& message : _messages)
                {
                    reply.append(std::format(" {}", message.uid));
                }
                reply.append(std::format("\r\n{} OK SEARCH completed\r\n", tag));
            }
            else if (verb == "UID FETCH")
            {
                const std::string_view set = std::string_view(argument).substr(0u, argument.find(' '));
                reply                      = BuildFetchRows(set);
                reply.append(std::format("{} OK FETCH completed\r\n", tag));
            }
            else if (verb == "NOOP")
            {
                reply = std::format("{} OK NOOP completed\r\n", tag);
            }
            else if (verb == "LOGOUT")
            {
                reply  = std::format("* BYE fake IMAP signing off\r\n{} OK LOGOUT completed\r\n", tag);
                logout = true;
            }
            else
            {
                reply = std::format("{} BAD unsupported fixture command\r\n", tag);
            }

            hr = SendAll(client, reply);
            if (FAILED(hr))
            {
                SetServerFailure(hr);
                return;
            }
            if (logout)
            {
                return;
            }
        }
    }

    mutable std::mutex _stateMutex;
    std::vector<FakeImapMessage> _messages;
    std::vector<CommandRecord> _commands;
    std::vector<std::string> _fetchSets;
    std::optional<uint64_t> _damagedUid;
    size_t _connections = 0u;
    size_t _logins      = 0u;
    HRESULT _serverHr   = S_OK;
    wil::unique_socket _listener;
    unsigned short _port = 0u;
    std::jthread _thread;
};
} // namespace

HRESULT ProbeCurlEphemeralBindForSelfTest(unsigned short& portOut) noexcept
{
    // Unlike the Common HTTP fixture or CreateLoopbackEndpoint, this observes
    // wildcard ephemeral allocation without listening, connecting, or reuse options.
    // The returned port is historical evidence, not a reservation for a later call.
    const int priorError = WSAGetLastError();
    auto restoreError    = wil::scope_exit([priorError]() noexcept { WSASetLastError(priorError); });
    portOut              = 0u;
    wil::unique_socket probe(socket(AF_INET, SOCK_STREAM, IPPROTO_TCP));
    if (! probe)
    {
        return SocketErrorToHResult();
    }
    sockaddr_in address{};
    address.sin_family      = AF_INET;
    address.sin_addr.s_addr = htonl(INADDR_ANY);
    if (bind(probe.get(), reinterpret_cast<const sockaddr*>(&address), static_cast<int>(sizeof(address))) != 0)
    {
        return SocketErrorToHResult();
    }
    int addressBytes = static_cast<int>(sizeof(address));
    if (getsockname(probe.get(), reinterpret_cast<sockaddr*>(&address), &addressBytes) != 0)
    {
        return SocketErrorToHResult();
    }
    portOut = ntohs(address.sin_port);
    return portOut != 0u ? S_OK : E_UNEXPECTED;
}

HRESULT ProbeCurlLoopbackConnectForSelfTest(bool bindFirst, unsigned short& localPortOut, const wchar_t*& stageOut) noexcept
{
    // Unlike the Common HTTP fixture, this owns only a TCP handshake: no worker,
    // accept, request or retry. Both modes use fresh listeners, not the failed peer.
    const int priorError = WSAGetLastError();
    auto restoreError    = wil::scope_exit([priorError]() noexcept { WSASetLastError(priorError); });
    localPortOut         = 0u;
    stageOut             = L"listener";
    wil::unique_socket listener;
    unsigned short listenerPort = 0u;
    const HRESULT listenerHr    = CreateLoopbackEndpoint(listener, listenerPort);
    if (FAILED(listenerHr))
    {
        return listenerHr;
    }

    stageOut = L"socket";
    wil::unique_socket probe(socket(AF_INET, SOCK_STREAM, IPPROTO_TCP));
    if (! probe)
    {
        return SocketErrorToHResult();
    }
    auto captureLocalPort = wil::scope_exit([&]() noexcept
    {
        sockaddr_in local{};
        int localBytes = static_cast<int>(sizeof(local));
        if (getsockname(probe.get(), reinterpret_cast<sockaddr*>(&local), &localBytes) == 0)
        {
            localPortOut = ntohs(local.sin_port);
        }
    });

    stageOut                  = L"nonblocking";
    unsigned long nonblocking = 1u;
    if (ioctlsocket(probe.get(), FIONBIO, &nonblocking) != 0)
    {
        return SocketErrorToHResult();
    }
    if (bindFirst)
    {
        stageOut = L"bind";
        sockaddr_in local{};
        local.sin_family      = AF_INET;
        local.sin_addr.s_addr = htonl(INADDR_ANY);
        if (bind(probe.get(), reinterpret_cast<const sockaddr*>(&local), static_cast<int>(sizeof(local))) != 0)
        {
            return SocketErrorToHResult();
        }
    }

    stageOut = L"connect";
    sockaddr_in peer{};
    peer.sin_family      = AF_INET;
    peer.sin_addr.s_addr = htonl(INADDR_LOOPBACK);
    peer.sin_port        = htons(listenerPort);
    if (connect(probe.get(), reinterpret_cast<const sockaddr*>(&peer), static_cast<int>(sizeof(peer))) != 0)
    {
        const int error = WSAGetLastError();
        if (error != WSAEWOULDBLOCK)
        {
            return HRESULT_FROM_WIN32(static_cast<unsigned long>(error));
        }
        stageOut = L"wait";
        fd_set writable;
        fd_set failed;
        FD_ZERO(&writable);
        FD_ZERO(&failed);
        FD_SET(probe.get(), &writable);
        FD_SET(probe.get(), &failed);
        timeval timeout{};
        timeout.tv_usec = 100000;
        const int ready = select(0, nullptr, &writable, &failed, &timeout);
        if (ready == SOCKET_ERROR)
        {
            return SocketErrorToHResult();
        }
        if (ready == 0)
        {
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }
        stageOut         = L"completion";
        int connectError = 0;
        int errorBytes   = static_cast<int>(sizeof(connectError));
        if (getsockopt(probe.get(), SOL_SOCKET, SO_ERROR, reinterpret_cast<char*>(&connectError), &errorBytes) != 0)
        {
            return SocketErrorToHResult();
        }
        if (connectError != 0)
        {
            return HRESULT_FROM_WIN32(static_cast<unsigned long>(connectError));
        }
        if (! FD_ISSET(probe.get(), &writable) || FD_ISSET(probe.get(), &failed))
        {
            return E_UNEXPECTED;
        }
    }
    stageOut = L"connected";
    return S_OK;
}

void RunCurlShortSuccessfulUploadSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    WSADATA winsockData{};
    const int startupResult = WSAStartup(MAKEWORD(2, 2), &winsockData);
    if (! DebugCheck(startupResult == 0, L"WSAStartup should initialize the loopback FTP proof", passed, failed))
    {
        return;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });

    try
    {
        const auto packageBCheck = [&](bool condition, const wchar_t* message) noexcept
        {
            if (! condition)
            {
                std::fputws(L"FileSystemCurl Package B selftest failed: ", stderr);
                std::fputws(message, stderr);
                std::fputwc(L'\n', stderr);
            }
            return DebugCheck(condition, message, passed, failed);
        };
        const auto packageCCheck = [&](bool condition, const wchar_t* message) noexcept
        {
            if (! condition)
            {
                std::fputws(L"FileSystemCurl Package C selftest failed: ", stderr);
                std::fputws(message, stderr);
                std::fputwc(L'\n', stderr);
            }
            return DebugCheck(condition, message, passed, failed);
        };

        const std::vector<uint8_t> sourceBytes         = MakeSourceBytes();
        const std::vector<uint8_t> destinationSentinel = MakeDestinationSentinel();
        constexpr std::string_view kConcurrentText      = "concurrent-owner-bytes-must-survive";
        const std::vector<uint8_t> concurrentBytes(kConcurrentText.begin(), kConcurrentText.end());
        const ScenarioResult shortResult               = RunScenario(UploadRetention::StrictPrefix, sourceBytes, destinationSentinel);
        CheckShortScenario(shortResult, sourceBytes, destinationSentinel, passed, failed);

        const ScenarioResult controlResult = RunScenario(UploadRetention::Complete, sourceBytes, destinationSentinel);
        CheckControlScenario(controlResult, sourceBytes, destinationSentinel, passed, failed);

        const WriterLifetimeResult writerLifetime =
            RunWriterScenario(sourceBytes, UploadRetention::Complete, &destinationSentinel, FILESYSTEM_FLAG_ALLOW_OVERWRITE, true);
        DebugCheck(writerLifetime.setupHr == S_OK && writerLifetime.createHr == S_OK && writerLifetime.writeHr == S_OK && writerLifetime.commitHr == S_OK &&
                       writerLifetime.written == sourceBytes.size(),
                   L"returned FTP writer should remain usable after every external filesystem interface is released",
                   passed,
                   failed);
        DebugCheck(FileEquals(writerLifetime.destination, "/owned-writer.bin", sourceBytes),
                   L"retained-owner FTP writer should commit the exact bytes before releasing its final owner",
                   passed,
                   failed);
        DebugCheck(TransactionSiblingEquals(writerLifetime.destination, kRollbackSiblingMarker, destinationSentinel) &&
                       CountCommands(writerLifetime.destination, "STOR", kWriterSiblingMarker) == 1u &&
                       CountCommands(writerLifetime.destination, "SIZE", kWriterSiblingMarker) == 1u &&
                       ! DeletedTransactionSibling(writerLifetime.destination, kRollbackSiblingMarker),
                   L"committed writer should stage, size-verify, promote, and preserve its overwrite backup without conditional identity",
                   passed,
                   failed);

        const WriterLifetimeResult abandonedWriter = RunWriterScenario(sourceBytes, UploadRetention::Complete, nullptr, FILESYSTEM_FLAG_ALLOW_OVERWRITE, false);
        DebugCheck(abandonedWriter.setupHr == S_OK && abandonedWriter.createHr == S_OK && abandonedWriter.writeHr == S_OK &&
                       abandonedWriter.commitHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && FileAbsent(abandonedWriter.destination, "/owned-writer.bin") &&
                       abandonedWriter.destination.commands.empty(),
                   L"abandoned writer should discard only its local staging file without contacting or mutating the endpoint",
                   passed,
                   failed);

        const WriterLifetimeResult shortWriter =
            RunWriterScenario(sourceBytes, UploadRetention::StrictPrefix, &destinationSentinel, FILESYSTEM_FLAG_ALLOW_OVERWRITE, true);
        DebugCheck(shortWriter.setupHr == S_OK && shortWriter.createHr == S_OK && shortWriter.writeHr == S_OK &&
                       shortWriter.commitHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                       FileEquals(shortWriter.destination, "/owned-writer.bin", destinationSentinel) && HasNoTransactionSiblings(shortWriter.destination) &&
                       DeletedTransactionSibling(shortWriter.destination, kWriterSiblingMarker),
                   L"short writer upload should remove staging and preserve the overwrite sentinel",
                   passed,
                   failed);

        const WriterLifetimeResult failedPromotion =
            RunWriterScenario(sourceBytes, UploadRetention::Complete, &destinationSentinel, FILESYSTEM_FLAG_ALLOW_OVERWRITE, true, true);
        DebugCheck(failedPromotion.setupHr == S_OK && failedPromotion.createHr == S_OK && failedPromotion.writeHr == S_OK &&
                       failedPromotion.commitHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                       FileAbsent(failedPromotion.destination, "/owned-writer.bin") &&
                       TransactionSiblingEquals(failedPromotion.destination, kRollbackSiblingMarker, destinationSentinel),
                   L"writer promotion failure should preserve the prior final as a rollback artifact rather than restore by pathname",
                   passed,
                   failed);

        const WriterLifetimeResult concurrentFailedPromotion = RunWriterScenario(
            sourceBytes, UploadRetention::Complete, &destinationSentinel, FILESYSTEM_FLAG_ALLOW_OVERWRITE, true, true, &concurrentBytes);
        packageCCheck(concurrentFailedPromotion.setupHr == S_OK && concurrentFailedPromotion.createHr == S_OK &&
                          concurrentFailedPromotion.writeHr == S_OK &&
                          concurrentFailedPromotion.commitHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                      L"writer promotion failure with a concurrent final should report partial preserved state");
        packageCCheck(FileEquals(concurrentFailedPromotion.destination, "/owned-writer.bin", concurrentBytes) &&
                          TransactionSiblingEquals(concurrentFailedPromotion.destination, kRollbackSiblingMarker, destinationSentinel) &&
                          concurrentFailedPromotion.destination.beforeRenameInjectionCount == 1u &&
                          CountCommands(concurrentFailedPromotion.destination, "DELE", "owned-writer.bin") == 0u,
                      L"writer promotion recovery should preserve the concurrent final and rollback backup without deleting by path");

        // R0f-Curl: a no-overwrite writer publishes new names (staged sibling + rename) and, when the
        // destination already exists, fails its Commit with ERROR_FILE_EXISTS leaving the existing
        // object and no staging sibling behind. The probe-to-rename window is documented, not refused.
        const WriterLifetimeResult noOverwriteExisting =
            RunWriterScenario(sourceBytes, UploadRetention::Complete, &destinationSentinel, FILESYSTEM_FLAG_NONE, true);
        DebugCheck(noOverwriteExisting.setupHr == S_OK && noOverwriteExisting.createHr == S_OK && noOverwriteExisting.writeHr == S_OK &&
                       noOverwriteExisting.commitHr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) &&
                       FileEquals(noOverwriteExisting.destination, "/owned-writer.bin", destinationSentinel) &&
                       HasNoTransactionSiblings(noOverwriteExisting.destination) &&
                       CountCommands(noOverwriteExisting.destination, "RNTO", "owned-writer.bin") == 0u,
                   L"writer no-overwrite publication onto an existing object must report ERROR_FILE_EXISTS, keep the object, and leave no staging sibling",
                   passed,
                   failed);
        const WriterLifetimeResult noOverwriteNew = RunWriterScenario(sourceBytes, UploadRetention::Complete, nullptr, FILESYSTEM_FLAG_NONE, true);
        DebugCheck(noOverwriteNew.setupHr == S_OK && noOverwriteNew.createHr == S_OK && noOverwriteNew.writeHr == S_OK && noOverwriteNew.commitHr == S_OK &&
                       FileEquals(noOverwriteNew.destination, "/owned-writer.bin", sourceBytes) && HasNoTransactionSiblings(noOverwriteNew.destination),
                   L"writer no-overwrite publication of a new name must publish exactly the staged bytes through one rename",
                   passed,
                   failed);

        uint64_t conditionalPublicationDurationUs = 0u;
        uint64_t conditionalPublicationCommands   = 0u;
        uint64_t conditionalPublicationInjections = 0u;
        const auto checkNoOverwriteSingle = [&](bool move, std::wstring_view operation) noexcept
        {
            const ScenarioResult result = RunNoOverwriteSingleRaceScenario(move, sourceBytes, concurrentBytes);
            conditionalPublicationDurationUs += result.elapsedMs * 1000u;
            conditionalPublicationCommands += result.source.commands.size() + result.destination.commands.size();
            conditionalPublicationInjections += result.destination.beforeRenameInjectionCount;
            // R0f-Curl: the destination is absent at probe time, so the staged bytes are published; the
            // creator injected between the probe and the rename is overwritten (documented window).
            packageCCheck(result.setupHr == S_OK && result.moveHr == S_OK &&
                              (move ? FileAbsent(result.source, "/source.bin") : FileEquals(result.source, "/source.bin", sourceBytes)) &&
                              FileEquals(result.destination, "/destination.bin", sourceBytes) &&
                              result.destination.beforeRenameInjectionCount == 1u,
                          std::format(L"no-overwrite single {} should probe, publish the staged bytes, and overwrite only the late creator", operation).c_str());
        };
        checkNoOverwriteSingle(false, L"COPY");
        checkNoOverwriteSingle(true, L"MOVE");

        std::vector<uint8_t> secondSourceBytes = sourceBytes;
        for (uint8_t& value : secondSourceBytes)
        {
            value ^= 0x5au;
        }
        const auto checkNoOverwriteBatch = [&](bool move, std::wstring_view operation) noexcept
        {
            const CollectionScenarioResult result = RunNoOverwriteBatchRaceScenario(move, sourceBytes, secondSourceBytes, concurrentBytes);
            conditionalPublicationDurationUs += result.elapsedMs * 1000u;
            conditionalPublicationCommands += result.source.commands.size() + result.destination.commands.size();
            conditionalPublicationInjections += result.destination.beforeRenameInjectionCount;
            packageCCheck(result.setupHr == S_OK && result.operationHr == S_OK &&
                              (move ? (FileAbsent(result.source, "/batch-a.bin") && FileAbsent(result.source, "/batch-b.bin"))
                                    : (FileEquals(result.source, "/batch-a.bin", sourceBytes) && FileEquals(result.source, "/batch-b.bin", secondSourceBytes))) &&
                              FileEquals(result.destination, "/batch-destination/batch-a.bin", sourceBytes) &&
                              FileEquals(result.destination, "/batch-destination/batch-b.bin", secondSourceBytes) &&
                              result.destination.beforeRenameInjectionCount == 2u,
                          std::format(L"no-overwrite batch {} should probe, publish every staged item, and overwrite only the late creators", operation).c_str());
        };
        checkNoOverwriteBatch(false, L"COPY");
        checkNoOverwriteBatch(true, L"MOVE");

        const auto checkNoOverwriteRename = [&](bool batch, std::wstring_view shape) noexcept
        {
            const RenameScenarioResult result = RunNoOverwriteRenameRaceScenario(batch, sourceBytes, concurrentBytes);
            conditionalPublicationDurationUs += result.elapsedMs * 1000u;
            conditionalPublicationCommands += result.endpoint.commands.size();
            conditionalPublicationInjections += result.endpoint.beforeRenameInjectionCount;
            const bool sourcesRenamed = FileAbsent(result.endpoint, "/rename-a.bin") && FileEquals(result.endpoint, "/renamed-a.bin", sourceBytes) &&
                                        (! batch || (FileAbsent(result.endpoint, "/rename-b.bin") && FileEquals(result.endpoint, "/renamed-b.bin", sourceBytes)));
            packageCCheck(result.setupHr == S_OK && result.operationHr == S_OK && sourcesRenamed &&
                              result.endpoint.beforeRenameInjectionCount == (batch ? 2u : 1u),
                          std::format(L"no-overwrite {} RENAME should probe, rename, and overwrite only a creator that appears inside the documented window", shape).c_str());
        };
        checkNoOverwriteRename(false, L"single");
        checkNoOverwriteRename(true, L"batch");

        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Curl.ConditionalPublicationGuard",
                              L"six direct no-overwrite FTP Copy/Move/Rename shapes; probe, publish, documented late-creator window",
                              conditionalPublicationDurationUs,
                              conditionalPublicationCommands,
                              conditionalPublicationInjections,
                              S_OK);
        }

        const ScenarioResult moveDeleteFailure =
            RunMoveDeleteFailureRaceScenario(sourceBytes, destinationSentinel, concurrentBytes);
        packageCCheck(moveDeleteFailure.setupHr == S_OK && moveDeleteFailure.moveHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                          FileEquals(moveDeleteFailure.source, "/source.bin", sourceBytes) && moveDeleteFailure.source.rejectedDeleteCount == 1u,
                      L"MOVE source-delete failure should preserve the source and report partial state");
        packageCCheck(FileEquals(moveDeleteFailure.destination, "/destination.bin", concurrentBytes) &&
                          TransactionSiblingEquals(moveDeleteFailure.destination, kRollbackSiblingMarker, destinationSentinel) &&
                          moveDeleteFailure.destination.afterRenameInjectionCount == 1u &&
                          CountCommands(moveDeleteFailure.destination, "DELE", "destination.bin") == 0u,
                      L"MOVE recovery should preserve a post-publication concurrent replacement and rollback backup without deleting by path");

        const CollectionScenarioResult recursiveRollback = RunRecursiveRollbackRaceScenario(sourceBytes, concurrentBytes);
        packageCCheck(recursiveRollback.setupHr == S_OK && recursiveRollback.operationHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                          FileEquals(recursiveRollback.source, "/source-tree/source.bin", sourceBytes),
                      L"failed recursive COPY should preserve its source and report partial state");
        packageCCheck(FileEquals(recursiveRollback.destination, "/destination-tree/concurrent.bin", concurrentBytes) &&
                          recursiveRollback.destination.afterDirectoryCreateInjectionCount == 1u &&
                          CountCommands(recursiveRollback.destination, "DELE", "concurrent.bin") == 0u &&
                          CountCommands(recursiveRollback.destination, "RMD", "destination-tree") == 0u,
                      L"recursive recovery should preserve a concurrent child and destination tree without pathname rollback");

        if (Debug::Perf::IsCaptureEnabled())
        {
            const uint64_t recoveryCommands = concurrentFailedPromotion.destination.commands.size() + moveDeleteFailure.source.commands.size() +
                                              moveDeleteFailure.destination.commands.size() + recursiveRollback.source.commands.size() +
                                              recursiveRollback.destination.commands.size();
            Debug::Perf::Emit(L"FileOps.Curl.IdentitySafeRecoveryGuard",
                              L"writer promotion failure, MOVE source-delete failure, and recursive COPY failure; preserve final/backup/tree",
                              (concurrentFailedPromotion.elapsedMs + moveDeleteFailure.elapsedMs + recursiveRollback.elapsedMs) * 1000u,
                              recoveryCommands,
                              3u,
                              HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY));
        }

        const auto packageDCheck = [&](bool condition, const wchar_t* message) noexcept
        {
            if (! condition)
            {
                std::fputws(L"FileSystemCurl Package D selftest failed: ", stderr);
                std::fputws(message, stderr);
                std::fputwc(L'\n', stderr);
            }
            return DebugCheck(condition, message, passed, failed);
        };

        const CleanupDebtScenarioResult cleanupDebtWriter =
            RunCleanupDebtWriterScenario(sourceBytes, destinationSentinel);
        packageDCheck(cleanupDebtWriter.setupHr == S_OK && cleanupDebtWriter.operationHr == S_OK &&
                          cleanupDebtWriter.secondCommitHr == S_OK && cleanupDebtWriter.callbacksSucceeded &&
                          FileEquals(cleanupDebtWriter.destination, "/owned-writer.bin", sourceBytes) &&
                          CountTransactionSiblings(cleanupDebtWriter.destination, kRollbackSiblingMarker) == 1u &&
                          TransactionSiblingEquals(cleanupDebtWriter.destination, kRollbackSiblingMarker, destinationSentinel) &&
                          ! DeletedTransactionSibling(cleanupDebtWriter.destination, kRollbackSiblingMarker) &&
                          CountCommands(cleanupDebtWriter.destination, "DELE", kRollbackSiblingMarker) == 0u &&
                          cleanupDebtWriter.destination.rejectedDeleteCount == 0u &&
                          CountCommands(cleanupDebtWriter.destination, "STOR", kWriterSiblingMarker) == 1u &&
                          CountCommands(cleanupDebtWriter.destination, "SIZE", kWriterSiblingMarker) == 1u &&
                          cleanupDebtWriter.commandsAfterFirstCommit == cleanupDebtWriter.destination.commands.size() &&
                          cleanupDebtWriter.alertsAfterFirstCommit == 1u && cleanupDebtWriter.alertCount == 1u &&
                          cleanupDebtWriter.notificationsAfterFirstCommit == 1u && cleanupDebtWriter.notificationCount == 1u &&
                          cleanupDebtWriter.alertShapeValid && CleanupDebtAlertIsContentFree(cleanupDebtWriter, 1u) &&
                          SUCCEEDED(cleanupDebtWriter.destination.serverHr) && cleanupDebtWriter.elapsedMs <= 15000u,
                      L"committed writer overwrite should expose one content-free cleanup debt exactly once and keep second Commit idempotent");

        const CleanupDebtScenarioResult failedWriterCleanup =
            RunCleanupDebtWriterScenario(sourceBytes, destinationSentinel, UploadRetention::StrictPrefix, true, false);
        packageDCheck(failedWriterCleanup.setupHr == S_OK &&
                          failedWriterCleanup.operationHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                          failedWriterCleanup.secondCommitHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && failedWriterCleanup.callbacksSucceeded &&
                          FileEquals(failedWriterCleanup.destination, "/owned-writer.bin", destinationSentinel) &&
                          CountTransactionSiblings(failedWriterCleanup.destination, kWriterSiblingMarker) == 1u &&
                          CountTransactionSiblings(failedWriterCleanup.destination, kRollbackSiblingMarker) == 0u &&
                          ! DeletedTransactionSibling(failedWriterCleanup.destination, kWriterSiblingMarker) &&
                          CountCommands(failedWriterCleanup.destination, "DELE", kWriterSiblingMarker) == 1u &&
                          failedWriterCleanup.destination.rejectedDeleteCount == 1u &&
                          failedWriterCleanup.commandsAfterFirstCommit == failedWriterCleanup.destination.commands.size() &&
                          failedWriterCleanup.alertsAfterFirstCommit == 1u && failedWriterCleanup.alertCount == 1u &&
                          failedWriterCleanup.notificationsAfterFirstCommit == 0u && failedWriterCleanup.notificationCount == 0u &&
                          failedWriterCleanup.alertShapeValid && CleanupDebtAlertIsContentFree(failedWriterCleanup, 1u) &&
                          SUCCEEDED(failedWriterCleanup.destination.serverHr) && failedWriterCleanup.elapsedMs <= 15000u,
                      L"failed writer verification should retain its primary error while exposing one failed staging-cleanup debt");

        const CleanupDebtScenarioResult failedWriterUpload =
            RunCleanupDebtWriterScenario(sourceBytes, destinationSentinel, UploadRetention::FailAfterStore, true, false);
        packageDCheck(failedWriterUpload.setupHr == S_OK && FAILED(failedWriterUpload.operationHr) &&
                          failedWriterUpload.operationHr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                          failedWriterUpload.secondCommitHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && failedWriterUpload.callbacksSucceeded &&
                          FileEquals(failedWriterUpload.destination, "/owned-writer.bin", destinationSentinel) &&
                          CountTransactionSiblings(failedWriterUpload.destination, kWriterSiblingMarker) == 1u &&
                          CountTransactionSiblings(failedWriterUpload.destination, kRollbackSiblingMarker) == 0u &&
                          ! DeletedTransactionSibling(failedWriterUpload.destination, kWriterSiblingMarker) &&
                          CountCommands(failedWriterUpload.destination, "STOR", kWriterSiblingMarker) == 2u &&
                          CountCommands(failedWriterUpload.destination, "SIZE", kWriterSiblingMarker) == 0u &&
                          CountCommands(failedWriterUpload.destination, "DELE", kWriterSiblingMarker) == 1u &&
                          failedWriterUpload.destination.rejectedDeleteCount == 1u &&
                          failedWriterUpload.commandsAfterFirstCommit == failedWriterUpload.destination.commands.size() &&
                          failedWriterUpload.alertsAfterFirstCommit == 1u && failedWriterUpload.alertCount == 1u &&
                          failedWriterUpload.notificationsAfterFirstCommit == 0u && failedWriterUpload.notificationCount == 0u &&
                          failedWriterUpload.alertShapeValid && CleanupDebtAlertIsContentFree(failedWriterUpload, 1u) &&
                          SUCCEEDED(failedWriterUpload.destination.serverHr) && failedWriterUpload.elapsedMs <= 15000u,
                      L"failed writer upload should retain its primary transport error and expose one failed staging-cleanup debt");

        const CleanupDebtScenarioResult cleanupDebtSingleCopy = RunCleanupDebtTransferScenario(
            CleanupDebtTransferShape::Single, false, sourceBytes, secondSourceBytes, destinationSentinel);
        packageDCheck(cleanupDebtSingleCopy.setupHr == S_OK && cleanupDebtSingleCopy.operationHr == S_OK &&
                          cleanupDebtSingleCopy.callbacksSucceeded && FileEquals(cleanupDebtSingleCopy.source, "/source.bin", sourceBytes) &&
                          FileEquals(cleanupDebtSingleCopy.destination, "/destination.bin", sourceBytes) &&
                          CountTransactionSiblings(cleanupDebtSingleCopy.destination, kRollbackSiblingMarker) == 1u &&
                          TransactionSiblingEquals(cleanupDebtSingleCopy.destination, kRollbackSiblingMarker, destinationSentinel) &&
                          ! DeletedTransactionSibling(cleanupDebtSingleCopy.destination, kRollbackSiblingMarker) &&
                          CountCommands(cleanupDebtSingleCopy.destination, "DELE", kRollbackSiblingMarker) == 0u &&
                          cleanupDebtSingleCopy.destination.rejectedDeleteCount == 0u &&
                          CountCommands(cleanupDebtSingleCopy.source, "DELE", "source.bin") == 0u &&
                          cleanupDebtSingleCopy.notificationCount == 1u && cleanupDebtSingleCopy.alertCount == 1u &&
                          cleanupDebtSingleCopy.alertShapeValid && CleanupDebtAlertIsContentFree(cleanupDebtSingleCopy, 1u) &&
                          SUCCEEDED(cleanupDebtSingleCopy.source.serverHr) && SUCCEEDED(cleanupDebtSingleCopy.destination.serverHr) &&
                          cleanupDebtSingleCopy.elapsedMs <= 15000u,
                      L"single COPY overwrite should commit bytes and expose one retained rollback debt without deleting it");

        const CleanupDebtScenarioResult cleanupDebtRecursiveCopy = RunCleanupDebtTransferScenario(
            CleanupDebtTransferShape::Recursive, false, sourceBytes, secondSourceBytes, destinationSentinel);
        packageDCheck(cleanupDebtRecursiveCopy.setupHr == S_OK && cleanupDebtRecursiveCopy.operationHr == S_OK &&
                          cleanupDebtRecursiveCopy.callbacksSucceeded &&
                          FileEquals(cleanupDebtRecursiveCopy.source, "/source-tree/root.bin", sourceBytes) &&
                          FileEquals(cleanupDebtRecursiveCopy.source, "/source-tree/nested/child.bin", secondSourceBytes) &&
                          FileEquals(cleanupDebtRecursiveCopy.destination, "/destination-tree/root.bin", sourceBytes) &&
                          FileEquals(cleanupDebtRecursiveCopy.destination, "/destination-tree/nested/child.bin", secondSourceBytes) &&
                          CountTransactionSiblings(cleanupDebtRecursiveCopy.destination, kRollbackSiblingMarker) == 2u &&
                          ! DeletedTransactionSibling(cleanupDebtRecursiveCopy.destination, kRollbackSiblingMarker) &&
                          CountCommands(cleanupDebtRecursiveCopy.destination, "DELE", kRollbackSiblingMarker) == 0u &&
                          cleanupDebtRecursiveCopy.destination.rejectedDeleteCount == 0u &&
                          cleanupDebtRecursiveCopy.notificationCount == 1u && cleanupDebtRecursiveCopy.alertCount == 1u &&
                          cleanupDebtRecursiveCopy.alertShapeValid && CleanupDebtAlertIsContentFree(cleanupDebtRecursiveCopy, 2u) &&
                          SUCCEEDED(cleanupDebtRecursiveCopy.source.serverHr) && SUCCEEDED(cleanupDebtRecursiveCopy.destination.serverHr) &&
                          cleanupDebtRecursiveCopy.elapsedMs <= 15000u,
                      L"recursive COPY overwrite should aggregate two retained rollback debts into one observable outcome");

        const CleanupDebtScenarioResult cleanupDebtBatchCopy = RunCleanupDebtTransferScenario(
            CleanupDebtTransferShape::Batch, false, sourceBytes, secondSourceBytes, destinationSentinel);
        packageDCheck(cleanupDebtBatchCopy.setupHr == S_OK && cleanupDebtBatchCopy.operationHr == S_OK &&
                          cleanupDebtBatchCopy.callbacksSucceeded && FileEquals(cleanupDebtBatchCopy.source, "/batch-a.bin", sourceBytes) &&
                          FileEquals(cleanupDebtBatchCopy.source, "/batch-b.bin", secondSourceBytes) &&
                          FileEquals(cleanupDebtBatchCopy.destination, "/batch-destination/batch-a.bin", sourceBytes) &&
                          FileEquals(cleanupDebtBatchCopy.destination, "/batch-destination/batch-b.bin", secondSourceBytes) &&
                          CountTransactionSiblings(cleanupDebtBatchCopy.destination, kRollbackSiblingMarker) == 2u &&
                          ! DeletedTransactionSibling(cleanupDebtBatchCopy.destination, kRollbackSiblingMarker) &&
                          CountCommands(cleanupDebtBatchCopy.destination, "DELE", kRollbackSiblingMarker) == 0u &&
                          cleanupDebtBatchCopy.destination.rejectedDeleteCount == 0u &&
                          cleanupDebtBatchCopy.notificationCount == 2u && cleanupDebtBatchCopy.alertCount == 1u &&
                          cleanupDebtBatchCopy.alertShapeValid && CleanupDebtAlertIsContentFree(cleanupDebtBatchCopy, 2u) &&
                          SUCCEEDED(cleanupDebtBatchCopy.source.serverHr) && SUCCEEDED(cleanupDebtBatchCopy.destination.serverHr) &&
                          cleanupDebtBatchCopy.elapsedMs <= 15000u,
                      L"batch COPY overwrite should aggregate two retained rollback debts while keeping both item callbacks successful");

        const CleanupDebtScenarioResult cleanupDebtSingleMove = RunCleanupDebtTransferScenario(
            CleanupDebtTransferShape::Single, true, sourceBytes, secondSourceBytes, destinationSentinel);
        packageDCheck(cleanupDebtSingleMove.setupHr == S_OK && cleanupDebtSingleMove.operationHr == S_OK &&
                          cleanupDebtSingleMove.callbacksSucceeded && FileAbsent(cleanupDebtSingleMove.source, "/source.bin") &&
                          FileEquals(cleanupDebtSingleMove.destination, "/destination.bin", sourceBytes) &&
                          CountTransactionSiblings(cleanupDebtSingleMove.destination, kRollbackSiblingMarker) == 1u &&
                          TransactionSiblingEquals(cleanupDebtSingleMove.destination, kRollbackSiblingMarker, destinationSentinel) &&
                          ! DeletedTransactionSibling(cleanupDebtSingleMove.destination, kRollbackSiblingMarker) &&
                          CountCommands(cleanupDebtSingleMove.destination, "DELE", kRollbackSiblingMarker) == 0u &&
                          cleanupDebtSingleMove.destination.rejectedDeleteCount == 0u &&
                          CountCommands(cleanupDebtSingleMove.source, "DELE", "source.bin") == 1u &&
                          cleanupDebtSingleMove.notificationCount == 1u && cleanupDebtSingleMove.alertCount == 1u &&
                          cleanupDebtSingleMove.alertShapeValid && CleanupDebtAlertIsContentFree(cleanupDebtSingleMove, 1u) &&
                          SUCCEEDED(cleanupDebtSingleMove.source.serverHr) && SUCCEEDED(cleanupDebtSingleMove.destination.serverHr) &&
                          cleanupDebtSingleMove.elapsedMs <= 15000u,
                      L"single MOVE overwrite should delete its source after publication and expose one retained rollback debt");

        const CleanupDebtScenarioResult cleanupDebtBatchMove = RunCleanupDebtTransferScenario(
            CleanupDebtTransferShape::Batch, true, sourceBytes, secondSourceBytes, destinationSentinel);
        packageDCheck(cleanupDebtBatchMove.setupHr == S_OK && cleanupDebtBatchMove.operationHr == S_OK &&
                          cleanupDebtBatchMove.callbacksSucceeded && FileAbsent(cleanupDebtBatchMove.source, "/batch-a.bin") &&
                          FileAbsent(cleanupDebtBatchMove.source, "/batch-b.bin") &&
                          FileEquals(cleanupDebtBatchMove.destination, "/batch-destination/batch-a.bin", sourceBytes) &&
                          FileEquals(cleanupDebtBatchMove.destination, "/batch-destination/batch-b.bin", secondSourceBytes) &&
                          CountTransactionSiblings(cleanupDebtBatchMove.destination, kRollbackSiblingMarker) == 2u &&
                          ! DeletedTransactionSibling(cleanupDebtBatchMove.destination, kRollbackSiblingMarker) &&
                          CountCommands(cleanupDebtBatchMove.destination, "DELE", kRollbackSiblingMarker) == 0u &&
                          cleanupDebtBatchMove.destination.rejectedDeleteCount == 0u &&
                          CountCommands(cleanupDebtBatchMove.source, "DELE", ".bin") == 2u &&
                          cleanupDebtBatchMove.notificationCount == 2u && cleanupDebtBatchMove.alertCount == 1u &&
                          cleanupDebtBatchMove.alertShapeValid && CleanupDebtAlertIsContentFree(cleanupDebtBatchMove, 2u) &&
                          SUCCEEDED(cleanupDebtBatchMove.source.serverHr) && SUCCEEDED(cleanupDebtBatchMove.destination.serverHr) &&
                          cleanupDebtBatchMove.elapsedMs <= 15000u,
                      L"batch MOVE overwrite should delete both sources and aggregate two retained rollback debts into one outcome");

        if (Debug::Perf::IsCaptureEnabled())
        {
            const uint64_t cleanupDebtDurationUs =
                (cleanupDebtWriter.elapsedMs + cleanupDebtSingleCopy.elapsedMs + cleanupDebtRecursiveCopy.elapsedMs +
                 cleanupDebtBatchCopy.elapsedMs + cleanupDebtSingleMove.elapsedMs + cleanupDebtBatchMove.elapsedMs) *
                1000u;
            const uint64_t cleanupDebtCommands = cleanupDebtWriter.destination.commands.size() + cleanupDebtSingleCopy.source.commands.size() +
                                                 cleanupDebtSingleCopy.destination.commands.size() + cleanupDebtRecursiveCopy.source.commands.size() +
                                                 cleanupDebtRecursiveCopy.destination.commands.size() + cleanupDebtBatchCopy.source.commands.size() +
                                                 cleanupDebtBatchCopy.destination.commands.size() + cleanupDebtSingleMove.source.commands.size() +
                                                 cleanupDebtSingleMove.destination.commands.size() + cleanupDebtBatchMove.source.commands.size() +
                                                 cleanupDebtBatchMove.destination.commands.size();
            Debug::Perf::Emit(L"FileOps.Curl.CleanupDebtGuard",
                              L"six committed overwrite shapes; aggregate retained recovery-item count and bounded endpoint work",
                              cleanupDebtDurationUs,
                              cleanupDebtCommands,
                              9u,
                              S_OK);
        }

        const ScenarioResult unknownSizeCopy = RunScenario(UploadRetention::StrictPrefix, sourceBytes, destinationSentinel, false, false);
        CheckUnknownSizeScenario(unknownSizeCopy, sourceBytes, destinationSentinel, L"COPY", passed, failed);
        const ScenarioResult unknownSizeMove = RunScenario(UploadRetention::StrictPrefix, sourceBytes, destinationSentinel, true, false);
        CheckUnknownSizeScenario(unknownSizeMove, sourceBytes, destinationSentinel, L"MOVE", passed, failed);

        const auto checkKnownSourceMismatch = [&](const ScenarioResult& result,
                                                  std::wstring_view operation,
                                                  std::wstring_view bodyShape) noexcept
        {
            packageBCheck(result.setupHr == S_OK && result.moveHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                          std::format(L"known-size {} RETR {} should fail with ERROR_PARTIAL_COPY (setup={:#010x}, result={:#010x})",
                                      bodyShape,
                                      operation,
                                      static_cast<unsigned long>(result.setupHr),
                                      static_cast<unsigned long>(result.moveHr))
                              .c_str());
            packageBCheck(FileEquals(result.source, "/source.bin", sourceBytes) &&
                           FileEquals(result.destination, "/destination.bin", destinationSentinel),
                          std::format(L"known-size {} RETR {} should preserve source and destination bytes", bodyShape, operation).c_str());
            packageBCheck(CountCommands(result.source, "RETR", "source.bin") == 1u &&
                           CountCommands(result.destination, "STOR") == 0u &&
                           CountCommands(result.source, "DELE", "source.bin") == 0u,
                          std::format(L"known-size {} RETR {} should stop before upload, promotion, or source deletion", bodyShape, operation).c_str());
        };

        const ScenarioResult knownSourceShortCopy =
            RunScenario(UploadRetention::Complete, sourceBytes, destinationSentinel, false, true, DownloadDelivery::StrictPrefix, true);
        checkKnownSourceMismatch(knownSourceShortCopy, L"COPY", L"short");
        const ScenarioResult knownSourceShortMove =
            RunScenario(UploadRetention::Complete, sourceBytes, destinationSentinel, true, true, DownloadDelivery::StrictPrefix, true);
        checkKnownSourceMismatch(knownSourceShortMove, L"MOVE", L"short");
        checkKnownSourceMismatch(
            RunScenario(UploadRetention::Complete, sourceBytes, destinationSentinel, false, true, DownloadDelivery::Overlong, true),
            L"COPY",
            L"overlong");
        checkKnownSourceMismatch(
            RunScenario(UploadRetention::Complete, sourceBytes, destinationSentinel, true, true, DownloadDelivery::Overlong, true),
            L"MOVE",
            L"overlong");

        const ScenarioResult unknownSourceMove =
            RunScenario(UploadRetention::Complete, sourceBytes, destinationSentinel, true, true, DownloadDelivery::Complete, false);
        packageBCheck(unknownSourceMove.setupHr == S_OK && unknownSourceMove.moveHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                       FileEquals(unknownSourceMove.source, "/source.bin", sourceBytes) &&
                       FileEquals(unknownSourceMove.destination, "/destination.bin", destinationSentinel),
                      L"unknown-size native MOVE should fail closed and preserve both endpoints");
        packageBCheck(CountCommands(unknownSourceMove.source, "RETR", "source.bin") == 0u &&
                       CountCommands(unknownSourceMove.destination, "STOR") == 0u &&
                       CountCommands(unknownSourceMove.source, "DELE", "source.bin") == 0u,
                      L"unknown-size native MOVE should stop before download, destination mutation, or source deletion");

        const ScenarioResult unknownSourceCopy =
            RunScenario(UploadRetention::Complete, sourceBytes, destinationSentinel, false, true, DownloadDelivery::Complete, false);
        packageBCheck(unknownSourceCopy.setupHr == S_OK && unknownSourceCopy.moveHr == S_OK &&
                       FileEquals(unknownSourceCopy.source, "/source.bin", sourceBytes) &&
                       FileEquals(unknownSourceCopy.destination, "/destination.bin", sourceBytes),
                      L"unknown-size native COPY may publish exact downloaded bytes while retaining its source");

        const auto checkKnownShortBatch = [&](const CollectionScenarioResult& result, std::wstring_view operation) noexcept
        {
            packageBCheck(result.setupHr == S_OK && result.operationHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                          std::format(L"known-size short RETR batch {} should report ERROR_PARTIAL_COPY (setup={:#010x}, result={:#010x})",
                                      operation,
                                      static_cast<unsigned long>(result.setupHr),
                                      static_cast<unsigned long>(result.operationHr))
                              .c_str());
            packageBCheck(FileEquals(result.source, "/batch-a.bin", sourceBytes) &&
                               FileEquals(result.source, "/batch-b.bin", secondSourceBytes) &&
                               FileAbsent(result.destination, "/batch-destination/batch-a.bin") &&
                               FileAbsent(result.destination, "/batch-destination/batch-b.bin"),
                          std::format(L"known-size short RETR batch {} should preserve both sources and publish neither destination", operation).c_str());
            packageBCheck(CountCommands(result.source, "RETR", ".bin") == 2u &&
                               CountCommands(result.destination, "STOR") == 0u &&
                               CountCommands(result.source, "DELE", ".bin") == 0u,
                          std::format(L"known-size short RETR batch {} should stop each item before upload or source deletion", operation).c_str());
        };

        checkKnownShortBatch(
            RunBatchScenario(false, sourceBytes, secondSourceBytes, DownloadDelivery::StrictPrefix, true), L"COPY");
        checkKnownShortBatch(
            RunBatchScenario(true, sourceBytes, secondSourceBytes, DownloadDelivery::StrictPrefix, true), L"MOVE");

        const auto checkKnownShortRecursive = [&](const CollectionScenarioResult& result, std::wstring_view operation) noexcept
        {
            packageBCheck(result.setupHr == S_OK && result.operationHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                          std::format(L"known-size short RETR recursive {} should report ERROR_PARTIAL_COPY (setup={:#010x}, result={:#010x})",
                                      operation,
                                      static_cast<unsigned long>(result.setupHr),
                                      static_cast<unsigned long>(result.operationHr))
                              .c_str());
            packageBCheck(FileEquals(result.source, "/source-tree/root.bin", sourceBytes) &&
                               FileEquals(result.source, "/source-tree/nested/child.bin", secondSourceBytes) &&
                               result.source.directories.contains("/source-tree") && result.source.directories.contains("/source-tree/nested") &&
                               result.destination.files.empty() && result.destination.directories.contains("/destination-tree") &&
                               CountCommands(result.destination, "RMD", "destination-tree") == 0u,
                          std::format(L"known-size short RETR recursive {} should preserve the source tree and the uncertain destination tree", operation)
                              .c_str());
            packageBCheck(CountCommands(result.source, "RETR", ".bin") == 1u &&
                               CountCommands(result.destination, "STOR") == 0u &&
                               CountCommands(result.source, "DELE", ".bin") == 0u,
                          std::format(L"known-size short RETR recursive {} should stop before upload or source deletion", operation).c_str());
        };

        checkKnownShortRecursive(
            RunRecursiveScenario(false, sourceBytes, secondSourceBytes, DownloadDelivery::StrictPrefix, true), L"COPY");
        checkKnownShortRecursive(
            RunRecursiveScenario(true, sourceBytes, secondSourceBytes, DownloadDelivery::StrictPrefix, true), L"MOVE");

        const CollectionScenarioResult unknownBatchCopy =
            RunBatchScenario(false, sourceBytes, secondSourceBytes, DownloadDelivery::Complete, false);
        packageBCheck(unknownBatchCopy.setupHr == S_OK && unknownBatchCopy.operationHr == S_OK &&
                           FileEquals(unknownBatchCopy.source, "/batch-a.bin", sourceBytes) &&
                           FileEquals(unknownBatchCopy.source, "/batch-b.bin", secondSourceBytes) &&
                           FileEquals(unknownBatchCopy.destination, "/batch-destination/batch-a.bin", sourceBytes) &&
                           FileEquals(unknownBatchCopy.destination, "/batch-destination/batch-b.bin", secondSourceBytes) &&
                           CountCommands(unknownBatchCopy.source, "DELE", ".bin") == 0u,
                      L"unknown-size batch COPY may publish both exact staged files while retaining both sources");

        const CollectionScenarioResult unknownBatchMove =
            RunBatchScenario(true, sourceBytes, secondSourceBytes, DownloadDelivery::Complete, false);
        packageBCheck(unknownBatchMove.setupHr == S_OK && unknownBatchMove.operationHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                           FileEquals(unknownBatchMove.source, "/batch-a.bin", sourceBytes) &&
                           FileEquals(unknownBatchMove.source, "/batch-b.bin", secondSourceBytes) && unknownBatchMove.destination.files.empty() &&
                           CountCommands(unknownBatchMove.source, "RETR", ".bin") == 0u &&
                           CountCommands(unknownBatchMove.destination, "STOR") == 0u &&
                           CountCommands(unknownBatchMove.source, "DELE", ".bin") == 0u,
                      L"unknown-size batch MOVE should fail every item before download, destination mutation, or source deletion");

        const CollectionScenarioResult unknownRecursiveCopy =
            RunRecursiveScenario(false, sourceBytes, secondSourceBytes, DownloadDelivery::Complete, false);
        packageBCheck(unknownRecursiveCopy.setupHr == S_OK && unknownRecursiveCopy.operationHr == S_OK &&
                           FileEquals(unknownRecursiveCopy.source, "/source-tree/root.bin", sourceBytes) &&
                           FileEquals(unknownRecursiveCopy.source, "/source-tree/nested/child.bin", secondSourceBytes) &&
                           FileEquals(unknownRecursiveCopy.destination, "/destination-tree/root.bin", sourceBytes) &&
                           FileEquals(unknownRecursiveCopy.destination, "/destination-tree/nested/child.bin", secondSourceBytes) &&
                           CountCommands(unknownRecursiveCopy.source, "DELE", ".bin") == 0u,
                      L"unknown-size recursive COPY may publish exact staged files while retaining the complete source tree");

        const CollectionScenarioResult unknownRecursiveMove =
            RunRecursiveScenario(true, sourceBytes, secondSourceBytes, DownloadDelivery::Complete, false);
        packageBCheck(unknownRecursiveMove.setupHr == S_OK && unknownRecursiveMove.operationHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                           FileEquals(unknownRecursiveMove.source, "/source-tree/root.bin", sourceBytes) &&
                           FileEquals(unknownRecursiveMove.source, "/source-tree/nested/child.bin", secondSourceBytes) &&
                           unknownRecursiveMove.destination.files.empty() && unknownRecursiveMove.destination.directories == std::set<std::string>{"/"} &&
                           CountCommands(unknownRecursiveMove.source, "RETR", ".bin") == 0u &&
                           CountCommands(unknownRecursiveMove.destination, "MKD") == 0u &&
                           CountCommands(unknownRecursiveMove.destination, "RMD") == 0u &&
                           CountCommands(unknownRecursiveMove.destination, "STOR") == 0u &&
                           CountCommands(unknownRecursiveMove.source, "DELE", ".bin") == 0u,
                      L"unknown-size recursive MOVE should fail preflight before creating a destination tree or touching source bytes");

        const auto checkHardProbeFailure = [&](const ScenarioResult& result, std::wstring_view operation) noexcept
        {
            packageBCheck(result.setupHr == S_OK && result.moveHr == HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE) &&
                               FileEquals(result.source, "/source.bin", sourceBytes) &&
                               FileEquals(result.destination, "/destination.bin", destinationSentinel) &&
                               CountCommands(result.source, "SIZE", "source.bin") >= 1u &&
                               CountCommands(result.source, "RETR", "source.bin") == 0u &&
                               CountCommands(result.destination, "STOR") == 0u &&
                               CountCommands(result.source, "DELE", "source.bin") == 0u,
                          std::format(L"targeted SIZE authentication failure for {} should propagate before transfer or mutation "
                                      L"(setup={:#010x}, result={:#010x}, SIZE={}, RETR={}, STOR={}, DELE={})",
                                      operation,
                                      static_cast<unsigned long>(result.setupHr),
                                      static_cast<unsigned long>(result.moveHr),
                                      CountCommands(result.source, "SIZE", "source.bin"),
                                      CountCommands(result.source, "RETR", "source.bin"),
                                      CountCommands(result.destination, "STOR"),
                                      CountCommands(result.source, "DELE", "source.bin"))
                              .c_str());
        };
        checkHardProbeFailure(RunScenario(UploadRetention::Complete,
                                          sourceBytes,
                                          destinationSentinel,
                                          false,
                                          true,
                                          DownloadDelivery::Complete,
                                          true,
                                          std::nullopt,
                                          true),
                              L"COPY");
        checkHardProbeFailure(RunScenario(UploadRetention::Complete,
                                          sourceBytes,
                                          destinationSentinel,
                                          true,
                                          true,
                                          DownloadDelivery::Complete,
                                          true,
                                          std::nullopt,
                                          true),
                              L"MOVE");

        const auto checkMissingProbeFailure = [&](const ScenarioResult& result, std::wstring_view operation) noexcept
        {
            packageBCheck(result.setupHr == S_OK && result.moveHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) &&
                               FileEquals(result.source, "/source.bin", sourceBytes) &&
                               FileEquals(result.destination, "/destination.bin", destinationSentinel) &&
                               CountCommands(result.source, "SIZE", "source.bin") >= 1u &&
                               CountCommands(result.source, "RETR", "source.bin") == 0u &&
                               CountCommands(result.destination, "STOR") == 0u &&
                               CountCommands(result.source, "DELE", "source.bin") == 0u,
                          std::format(L"targeted SIZE disappearance for {} should propagate before transfer or mutation "
                                      L"(setup={:#010x}, result={:#010x}, SIZE={}, RETR={}, STOR={}, DELE={})",
                                      operation,
                                      static_cast<unsigned long>(result.setupHr),
                                      static_cast<unsigned long>(result.moveHr),
                                      CountCommands(result.source, "SIZE", "source.bin"),
                                      CountCommands(result.source, "RETR", "source.bin"),
                                      CountCommands(result.destination, "STOR"),
                                      CountCommands(result.source, "DELE", "source.bin"))
                              .c_str());
        };
        checkMissingProbeFailure(RunScenario(UploadRetention::Complete,
                                             sourceBytes,
                                             destinationSentinel,
                                             false,
                                             true,
                                             DownloadDelivery::Complete,
                                             true,
                                             std::nullopt,
                                             false,
                                             true),
                                 L"COPY");
        checkMissingProbeFailure(RunScenario(UploadRetention::Complete,
                                             sourceBytes,
                                             destinationSentinel,
                                             true,
                                             true,
                                             DownloadDelivery::Complete,
                                             true,
                                             std::nullopt,
                                             false,
                                             true),
                                 L"MOVE");

        const auto checkStaleSingle = [&](const ScenarioResult& result, bool move, std::wstring_view operation) noexcept
        {
            packageBCheck(result.setupHr == S_OK && result.moveHr == S_OK &&
                               (move ? FileAbsent(result.source, "/source.bin") : FileEquals(result.source, "/source.bin", sourceBytes)) &&
                               FileEquals(result.destination, "/destination.bin", sourceBytes) &&
                               CountCommands(result.source, "SIZE", "source.bin") >= 1u,
                          std::format(L"stale-listing single {} should use the targeted source size and publish exact bytes", operation).c_str());
        };
        checkStaleSingle(RunScenario(UploadRetention::Complete,
                                     sourceBytes,
                                     destinationSentinel,
                                     false,
                                     true,
                                     DownloadDelivery::Complete,
                                     true,
                                     sourceBytes.size() / 2u),
                         false,
                         L"COPY");
        checkStaleSingle(RunScenario(UploadRetention::Complete,
                                     sourceBytes,
                                     destinationSentinel,
                                     true,
                                     true,
                                     DownloadDelivery::Complete,
                                     true,
                                     sourceBytes.size() / 2u),
                         true,
                         L"MOVE");

        const auto checkStaleBatch = [&](const CollectionScenarioResult& result, bool move, std::wstring_view operation) noexcept
        {
            const bool sourceState = move ? FileAbsent(result.source, "/batch-a.bin") && FileAbsent(result.source, "/batch-b.bin")
                                          : FileEquals(result.source, "/batch-a.bin", sourceBytes) &&
                                                FileEquals(result.source, "/batch-b.bin", secondSourceBytes);
            packageBCheck(result.setupHr == S_OK && result.operationHr == S_OK && sourceState &&
                               FileEquals(result.destination, "/batch-destination/batch-a.bin", sourceBytes) &&
                               FileEquals(result.destination, "/batch-destination/batch-b.bin", secondSourceBytes) &&
                               CountCommands(result.source, "SIZE", ".bin") >= 2u,
                          std::format(L"stale-listing batch {} should target-probe both files and publish exact bytes", operation).c_str());
        };
        checkStaleBatch(RunBatchScenario(false, sourceBytes, secondSourceBytes, DownloadDelivery::Complete, true, true), false, L"COPY");
        checkStaleBatch(RunBatchScenario(true, sourceBytes, secondSourceBytes, DownloadDelivery::Complete, true, true), true, L"MOVE");

        const auto checkStaleRecursive = [&](const CollectionScenarioResult& result, bool move, std::wstring_view operation) noexcept
        {
            const bool sourceState = move ? FileAbsent(result.source, "/source-tree/root.bin") &&
                                                FileAbsent(result.source, "/source-tree/nested/child.bin") &&
                                                ! result.source.directories.contains("/source-tree")
                                          : FileEquals(result.source, "/source-tree/root.bin", sourceBytes) &&
                                                FileEquals(result.source, "/source-tree/nested/child.bin", secondSourceBytes);
            packageBCheck(result.setupHr == S_OK && result.operationHr == S_OK && sourceState &&
                               FileEquals(result.destination, "/destination-tree/root.bin", sourceBytes) &&
                               FileEquals(result.destination, "/destination-tree/nested/child.bin", secondSourceBytes) &&
                               CountCommands(result.source, "SIZE", ".bin") >= 2u,
                          std::format(L"stale-listing recursive {} should target-probe every file and publish the exact tree", operation).c_str());
        };
        checkStaleRecursive(
            RunRecursiveScenario(false, sourceBytes, secondSourceBytes, DownloadDelivery::Complete, true, true), false, L"COPY");
        checkStaleRecursive(
            RunRecursiveScenario(true, sourceBytes, secondSourceBytes, DownloadDelivery::Complete, true, true), true, L"MOVE");

        const ReaderScenarioResult exactReader = RunReaderScenario(sourceBytes, DownloadDelivery::Complete);
        packageBCheck(exactReader.setupHr == S_OK && exactReader.createHr == S_OK && exactReader.getSizeHr == S_OK &&
                       exactReader.committedSizeBytes == sourceBytes.size() && exactReader.firstReadHr == S_OK &&
                       exactReader.firstBytesRead == sourceBytes.size() && exactReader.firstBytes == sourceBytes &&
                       exactReader.secondReadHr == S_OK && exactReader.secondBytesRead == 0u,
                      L"real Curl reader should expose exact known-size bytes followed by successful EOF");

        const std::vector<uint8_t> emptySourceBytes;
        const ReaderScenarioResult exactEmptyReader = RunReaderScenario(emptySourceBytes, DownloadDelivery::Complete);
        packageBCheck(exactEmptyReader.setupHr == S_OK && exactEmptyReader.createHr == S_OK && exactEmptyReader.getSizeHr == S_OK &&
                           exactEmptyReader.committedSizeBytes == 0u && exactEmptyReader.firstReadHr == S_OK &&
                           exactEmptyReader.firstBytesRead == 0u && exactEmptyReader.secondReadHr == S_OK &&
                           exactEmptyReader.secondBytesRead == 0u && CountCommands(exactEmptyReader.source, "RETR", "source.bin") == 1u,
                      L"real Curl reader should validate an exact zero-byte body before returning successful EOF");

        constexpr unsigned int kReaderShutdownStressIterations = 32u;
        bool readerShutdownStressPassed                         = true;
        for (unsigned int iteration = 0u; iteration < kReaderShutdownStressIterations; ++iteration)
        {
            const ReaderScenarioResult shutdownReader = RunReaderScenario(emptySourceBytes, DownloadDelivery::Complete);
            readerShutdownStressPassed =
                readerShutdownStressPassed && shutdownReader.setupHr == S_OK && shutdownReader.createHr == S_OK && shutdownReader.getSizeHr == S_OK &&
                shutdownReader.committedSizeBytes == 0u && shutdownReader.firstReadHr == S_OK && shutdownReader.firstBytesRead == 0u &&
                shutdownReader.secondReadHr == S_OK && shutdownReader.secondBytesRead == 0u &&
                CountCommands(shutdownReader.source, "RETR", "source.bin") == 1u;
        }
        packageBCheck(readerShutdownStressPassed,
                      L"Curl reader teardown should wake and join an idle transfer worker without a lost notification");

        const ReaderScenarioResult overlongEmptyReader = RunReaderScenario(emptySourceBytes, DownloadDelivery::Overlong);
        packageBCheck(overlongEmptyReader.setupHr == S_OK && overlongEmptyReader.createHr == S_OK && overlongEmptyReader.getSizeHr == S_OK &&
                           overlongEmptyReader.committedSizeBytes == 0u &&
                           overlongEmptyReader.firstReadHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) &&
                           overlongEmptyReader.firstBytesRead == 0u && overlongEmptyReader.source.lastDownloadSentBytes == 1u,
                      L"real Curl reader should reject a nonempty body that contradicts a zero-byte commitment");

        const ReaderScenarioResult staleReader =
            RunReaderScenario(sourceBytes, DownloadDelivery::Complete, std::nullopt, false, sourceBytes.size() / 2u);
        packageBCheck(staleReader.setupHr == S_OK && staleReader.createHr == S_OK && staleReader.getSizeHr == S_OK &&
                           staleReader.committedSizeBytes == sourceBytes.size() && staleReader.firstReadHr == S_OK &&
                           staleReader.firstBytes == sourceBytes && staleReader.secondReadHr == S_OK && staleReader.secondBytesRead == 0u &&
                           CountCommands(staleReader.source, "SIZE", "source.bin") >= 1u,
                      L"real Curl reader should replace a stale listing size with the targeted source commitment");

        const ReaderScenarioResult readerProbeAuthentication =
            RunReaderScenario(sourceBytes, DownloadDelivery::Complete, std::nullopt, false, std::nullopt, true);
        packageBCheck(readerProbeAuthentication.setupHr == S_OK &&
                           readerProbeAuthentication.createHr == HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE) &&
                           CountCommands(readerProbeAuthentication.source, "SIZE", "source.bin") >= 1u &&
                           CountCommands(readerProbeAuthentication.source, "RETR", "source.bin") == 0u,
                      std::format(L"real Curl reader should propagate a targeted SIZE authentication failure before starting RETR "
                                  L"(setup={:#010x}, create={:#010x}, SIZE={}, RETR={})",
                                  static_cast<unsigned long>(readerProbeAuthentication.setupHr),
                                  static_cast<unsigned long>(readerProbeAuthentication.createHr),
                                  CountCommands(readerProbeAuthentication.source, "SIZE", "source.bin"),
                                  CountCommands(readerProbeAuthentication.source, "RETR", "source.bin"))
                          .c_str());

        const ReaderScenarioResult readerProbeMissing =
            RunReaderScenario(sourceBytes, DownloadDelivery::Complete, std::nullopt, false, std::nullopt, false, true);
        packageBCheck(readerProbeMissing.setupHr == S_OK && readerProbeMissing.createHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) &&
                           CountCommands(readerProbeMissing.source, "SIZE", "source.bin") >= 1u &&
                           CountCommands(readerProbeMissing.source, "RETR", "source.bin") == 0u,
                      std::format(L"real Curl reader should propagate targeted SIZE disappearance before RETR "
                                  L"(setup={:#010x}, create={:#010x}, SIZE={}, RETR={})",
                                  static_cast<unsigned long>(readerProbeMissing.setupHr),
                                  static_cast<unsigned long>(readerProbeMissing.createHr),
                                  CountCommands(readerProbeMissing.source, "SIZE", "source.bin"),
                                  CountCommands(readerProbeMissing.source, "RETR", "source.bin"))
                          .c_str());

        const ReaderScenarioResult shortReader = RunReaderScenario(sourceBytes, DownloadDelivery::StrictPrefix);
        packageBCheck(shortReader.setupHr == S_OK && shortReader.createHr == S_OK && shortReader.getSizeHr == S_OK &&
                       shortReader.firstReadHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) && shortReader.firstBytesRead == 0u,
                      std::format(L"real Curl reader should reject protocol-success premature EOF before returning successful zero bytes "
                                  L"(setup={:#010x}, create={:#010x}, size={:#010x}, read={:#010x}, bytes={})",
                                  static_cast<unsigned long>(shortReader.setupHr),
                                  static_cast<unsigned long>(shortReader.createHr),
                                  static_cast<unsigned long>(shortReader.getSizeHr),
                                  static_cast<unsigned long>(shortReader.firstReadHr),
                                  shortReader.firstBytesRead)
                          .c_str());

        const ReaderScenarioResult overlongReader = RunReaderScenario(sourceBytes, DownloadDelivery::Overlong);
        packageBCheck(overlongReader.setupHr == S_OK && overlongReader.createHr == S_OK && overlongReader.getSizeHr == S_OK &&
                       overlongReader.firstReadHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && overlongReader.firstBytesRead == 0u,
                      std::format(L"real Curl reader should reject a body beyond its committed size "
                                  L"(setup={:#010x}, create={:#010x}, size={:#010x}, read={:#010x}, bytes={})",
                                  static_cast<unsigned long>(overlongReader.setupHr),
                                  static_cast<unsigned long>(overlongReader.createHr),
                                  static_cast<unsigned long>(overlongReader.getSizeHr),
                                  static_cast<unsigned long>(overlongReader.firstReadHr),
                                  overlongReader.firstBytesRead)
                          .c_str());

        constexpr uint64_t kSeekOffset = 137u;
        const ReaderScenarioResult seekReader = RunReaderScenario(sourceBytes, DownloadDelivery::Complete, kSeekOffset);
        const std::vector<uint8_t> expectedSeekBytes(sourceBytes.begin() + static_cast<std::ptrdiff_t>(kSeekOffset), sourceBytes.end());
        packageBCheck(seekReader.setupHr == S_OK && seekReader.createHr == S_OK && seekReader.getSizeHr == S_OK &&
                       seekReader.committedSizeBytes == sourceBytes.size() && seekReader.seekHr == S_OK &&
                       seekReader.seekPosition == kSeekOffset && seekReader.firstReadHr == S_OK &&
                       seekReader.firstBytes == expectedSeekBytes && seekReader.secondReadHr == S_OK && seekReader.secondBytesRead == 0u &&
                       seekReader.source.lastDownloadSentBytes == expectedSeekBytes.size() &&
                       CountCommands(seekReader.source, "REST", "137") == 1u &&
                       CountCommands(seekReader.source, "RETR", "source.bin") == 1u,
                      (std::format(L"real Curl reader should validate a REST seek against only the committed remaining range "
                                   L"(setup={:#010x}, create={:#010x}, seek={:#010x}/{}, read={:#010x}/{}, eof={:#010x}/{}, expected={})",
                                   static_cast<unsigned long>(seekReader.setupHr),
                                   static_cast<unsigned long>(seekReader.createHr),
                                   static_cast<unsigned long>(seekReader.seekHr),
                                   seekReader.seekPosition,
                                   static_cast<unsigned long>(seekReader.firstReadHr),
                                   seekReader.firstBytesRead,
                                   static_cast<unsigned long>(seekReader.secondReadHr),
                                   seekReader.secondBytesRead,
                                   expectedSeekBytes.size()) +
                       std::format(L", sent={}, REST={}, RETR={}",
                                   seekReader.source.lastDownloadSentBytes,
                                   CountCommands(seekReader.source, "REST", "137"),
                                   CountCommands(seekReader.source, "RETR", "source.bin")))
                          .c_str());

        const ReaderScenarioResult restartReader = RunReaderScenario(sourceBytes, DownloadDelivery::StrictPrefix, std::nullopt, true);
        packageBCheck(restartReader.firstReadHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) && restartReader.restartSeekHr == S_OK &&
                       restartReader.restartReadHr == S_OK && restartReader.restartBytes == sourceBytes &&
                       restartReader.restartEofHr == S_OK && restartReader.restartEofBytes == 0u,
                      std::format(L"real Curl reader should discard a failed generation and validate a complete restarted generation "
                                  L"(first={:#010x}, seek={:#010x}, restart={:#010x}/{}, eof={:#010x}/{})",
                                  static_cast<unsigned long>(restartReader.firstReadHr),
                                  static_cast<unsigned long>(restartReader.restartSeekHr),
                                  static_cast<unsigned long>(restartReader.restartReadHr),
                                  restartReader.restartBytesRead,
                                  static_cast<unsigned long>(restartReader.restartEofHr),
                                  restartReader.restartEofBytes)
                          .c_str());

        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Curl.ShortSuccessfulUpload",
                              L"real loopback FTP MOVE; complete stream consumed; strict prefix retained; protocol 226 success",
                              shortResult.elapsedMs * 1000u,
                              shortResult.destination.commands.size(),
                              shortResult.destination.lastUploadRetainedBytes,
                              shortResult.moveHr);
            Debug::Perf::Emit(L"FileOps.Curl.CompleteUploadControl",
                              L"real loopback FTP MOVE; complete stream retained and promoted over an existing destination",
                              controlResult.elapsedMs * 1000u,
                              controlResult.destination.commands.size(),
                              controlResult.destination.lastUploadRetainedBytes,
                              controlResult.moveHr);
            Debug::Perf::Emit(L"FileOps.Curl.UnknownSizeFailClosed",
                              L"real loopback FTP COPY/MOVE; SIZE rejected; LIST omits size; staged upload cleaned",
                              unknownSizeCopy.elapsedMs * 1000u + unknownSizeMove.elapsedMs * 1000u,
                              unknownSizeCopy.destination.commands.size() + unknownSizeMove.destination.commands.size(),
                              unknownSizeCopy.destination.lastUploadRetainedBytes + unknownSizeMove.destination.lastUploadRetainedBytes,
                              unknownSizeMove.moveHr);
            Debug::Perf::Emit(L"FileOps.Curl.SourceCommitmentGuard",
                              L"real loopback FTP COPY/MOVE; targeted SIZE commitment; clean short RETR stops before STOR/DELE",
                              knownSourceShortCopy.elapsedMs * 1000u + knownSourceShortMove.elapsedMs * 1000u,
                              knownSourceShortCopy.source.commands.size() + knownSourceShortMove.source.commands.size(),
                              knownSourceShortCopy.source.lastDownloadSentBytes + knownSourceShortMove.source.lastDownloadSentBytes,
                              knownSourceShortMove.moveHr);
            Debug::Perf::Emit(L"FileOps.Curl.Writer.TransactionalCommit",
                              L"real loopback FTP writer; local staging; exact remote size; overwrite backup and promotion",
                              writerLifetime.elapsedMs * 1000u,
                              writerLifetime.destination.commands.size(),
                              writerLifetime.written,
                              writerLifetime.commitHr);
            Debug::Perf::Emit(L"FileOps.Curl.Writer.ShortUploadCleanup",
                              L"real loopback FTP writer; protocol success retains a strict prefix; final sentinel preserved",
                              shortWriter.elapsedMs * 1000u,
                              shortWriter.destination.commands.size(),
                              shortWriter.destination.lastUploadRetainedBytes,
                              shortWriter.commitHr);
            Debug::Perf::Emit(L"FileOps.Curl.Writer.PromotionRollback",
                              L"real loopback FTP writer; staged upload succeeds; promotion fails; overwrite backup preserved",
                              failedPromotion.elapsedMs * 1000u,
                              failedPromotion.destination.commands.size(),
                              failedPromotion.written,
                              failedPromotion.commitHr);
            Debug::Perf::Emit(L"FileOps.Curl.Writer.Abandoned",
                              L"writer released before Commit; local staging discarded; zero endpoint commands",
                              abandonedWriter.elapsedMs * 1000u,
                              abandonedWriter.destination.commands.size(),
                              abandonedWriter.written,
                              abandonedWriter.commitHr);
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported-selftest boundary: report deterministic failure instead of escaping through the C ABI.
        Debug::Error(L"FileSystemCurl FTP transport selftest failed after std::exception.");
        DebugCheck(false, L"loopback FTP proof should not throw std::exception", passed, failed);
    }
}
// R0f-Curl witness: a control command the server never answers returns to the caller within about
// a second of Cancel (progress-callback polling), and on its own within the provider-owned bound
// when nobody cancels (transport response timeout).
// R0f-Curl-OR2: several threads upload through one plugin instance at the same time, as the host's
// parallel copy does. Every easy handle keeps its own libcurl connections (the process-wide share
// covers DNS and TLS sessions only), so concurrent transfers never touch a shared connection pool.
// Three Fresh Full gates crashed inside libcurl's pooled-connection matching before this held.
void RunCurlParallelWritersSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    constexpr unsigned int kThreads         = 6u;
    constexpr unsigned int kUploadsPerThread = 10u;
    constexpr size_t kUploadBytes            = 64u * 1024u;

    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"WSAStartup should initialize the parallel-writers proof", passed, failed))
    {
        return;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });

    try
    {
        FakeFtpEndpoint endpoint(UploadRetention::Complete);
        if (! DebugCheck(SUCCEEDED(endpoint.Start()), L"parallel-writers fixture should start", passed, failed))
        {
            return;
        }
        auto stopEndpoint = wil::scope_exit([&]() noexcept { endpoint.Stop(); });

        wil::com_ptr<IFileSystem> fileSystem;
        HRESULT hr = CreateFtpSelfTestInstance(fileSystem.put());
        if (! DebugCheck(SUCCEEDED(hr) && fileSystem, L"parallel-writers fixture should create an FTP instance", passed, failed))
        {
            return;
        }
        wil::com_ptr<IInformations> information;
        hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
        if (! DebugCheck(SUCCEEDED(hr) && information, L"FTP instance should expose IInformations", passed, failed))
        {
            return;
        }
        hr = information->SetConfiguration(R"({"connectTimeoutMs":5000,"operationTimeoutMs":30000,"copyMoveMaxConcurrency":8,"deleteMaxConcurrency":1,"ftpUseEpsv":true})");
        if (! DebugCheck(SUCCEEDED(hr), L"FTP instance should accept the parallel-writers configuration", passed, failed))
        {
            return;
        }
        wil::com_ptr<IFileSystemIO> io;
        hr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), io.put_void());
        if (! DebugCheck(SUCCEEDED(hr) && io, L"FTP instance should expose IFileSystemIO", passed, failed))
        {
            return;
        }

        const std::vector<uint8_t> payload(kUploadBytes, 0xA7u);
        std::atomic<unsigned int> failures{0u};
        std::atomic<unsigned int> uploads{0u};
        {
            std::vector<std::jthread> workers;
            workers.reserve(kThreads);
            for (unsigned int thread = 0u; thread < kThreads; ++thread)
            {
                workers.emplace_back([&, thread]() noexcept
                {
                    for (unsigned int index = 0u; index < kUploadsPerThread; ++index)
                    {
                        const std::wstring path =
                            std::format(L"//anonymous@127.0.0.1:{}/parallel-{}-{}.bin", endpoint.Port(), thread, index);
                        wil::com_ptr<IFileWriter> writer;
                        HRESULT uploadHr = io->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.put());
                        unsigned long written = 0u;
                        if (SUCCEEDED(uploadHr) && writer)
                        {
                            uploadHr = writer->Write(payload.data(), static_cast<unsigned long>(payload.size()), &written);
                        }
                        if (SUCCEEDED(uploadHr) && writer)
                        {
                            uploadHr = writer->Commit();
                        }
                        if (FAILED(uploadHr) || written != payload.size())
                        {
                            failures.fetch_add(1u, std::memory_order_acq_rel);
                        }
                        else
                        {
                            uploads.fetch_add(1u, std::memory_order_acq_rel);
                        }
                    }
                });
            }
        }

        DebugCheck(failures.load(std::memory_order_acquire) == 0u && uploads.load(std::memory_order_acquire) == kThreads * kUploadsPerThread,
                   L"every concurrent upload through one FTP instance must complete (per-handle connections, no shared pool)",
                   passed,
                   failed);
        DebugCheck(CountCommands(endpoint.Snapshot(), "STOR") >= kThreads * kUploadsPerThread,
                   L"the fake FTP server should have received every concurrent STOR",
                   passed,
                   failed);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported-selftest boundary: record the failure instead of escaping through C ABI.
        Debug::Error(L"FileSystemCurl parallel-writers selftest failed after std::exception.");
        DebugCheck(false, L"parallel-writers proof should not throw std::exception", passed, failed);
    }
}

// R0f-Curl-OR1: a Read that is waiting inside libcurl (stalled RETR) ends with ERROR_CANCELLED when the
// operation control handed to the reader cancels, without a host stop.
void RunCurlStalledReaderCancelSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    class ReaderCancelControl final : public IFileSystemOperationControl
    {
    public:
        ReaderCancelControl() noexcept = default;
        ReaderCancelControl(const ReaderCancelControl&)            = delete;
        ReaderCancelControl& operator=(const ReaderCancelControl&) = delete;
        ReaderCancelControl(ReaderCancelControl&&)                 = delete;
        ReaderCancelControl& operator=(ReaderCancelControl&&)      = delete;

        HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void*) noexcept override
        {
            if (abort == nullptr)
            {
                return E_POINTER;
            }
            *abort = abortRequested.load(std::memory_order_acquire) ? TRUE : FALSE;
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
        std::atomic<bool> abortRequested{false};
    };

    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"WSAStartup should initialize the stalled-reader proof", passed, failed))
    {
        return;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });

    try
    {
        FakeFtpEndpoint endpoint(UploadRetention::Complete);
        endpoint.SeedFile("/stall-read.bin", std::vector<uint8_t>(256u * 1024u, 0x5Au));
        endpoint.SetStallVerb("RETR", 30'000u);
        if (! DebugCheck(SUCCEEDED(endpoint.Start()), L"stalled-reader fixture should start", passed, failed))
        {
            return;
        }
        auto stopEndpoint = wil::scope_exit([&]() noexcept { endpoint.Stop(); });

        wil::com_ptr<IFileSystem> fileSystem;
        HRESULT hr = CreateFtpSelfTestInstance(fileSystem.put());
        if (! DebugCheck(SUCCEEDED(hr) && fileSystem, L"stalled-reader fixture should create an FTP instance", passed, failed))
        {
            return;
        }
        wil::com_ptr<IInformations> information;
        hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
        if (! DebugCheck(SUCCEEDED(hr) && information, L"FTP instance should expose IInformations", passed, failed))
        {
            return;
        }
        // A generous watchdog: the cancel, not the low-speed bound, must end the stalled transfer.
        hr = information->SetConfiguration(R"({"connectTimeoutMs":5000,"operationTimeoutMs":20000,"copyMoveMaxConcurrency":1,"deleteMaxConcurrency":1,"ftpUseEpsv":true})");
        if (! DebugCheck(SUCCEEDED(hr), L"FTP instance should accept the stalled-reader configuration", passed, failed))
        {
            return;
        }
        wil::com_ptr<IFileSystemIO> io;
        hr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), io.put_void());
        if (! DebugCheck(SUCCEEDED(hr) && io, L"FTP instance should expose IFileSystemIO", passed, failed))
        {
            return;
        }

        const std::wstring stalledPath = std::format(L"//anonymous@127.0.0.1:{}/stall-read.bin", endpoint.Port());
        wil::com_ptr<IFileReader> reader;
        hr = io->CreateFileReader(stalledPath.c_str(), reader.put());
        if (! DebugCheck(SUCCEEDED(hr) && reader, L"CreateFileReader should succeed before the stalled RETR", passed, failed))
        {
            return;
        }
        wil::com_ptr<IFileReaderOperationControl> readerControl;
        hr = reader->QueryInterface(IID_PPV_ARGS(readerControl.addressof()));
        if (! DebugCheck(SUCCEEDED(hr) && readerControl, L"the streaming reader must expose IFileReaderOperationControl", passed, failed))
        {
            return;
        }

        ReaderCancelControl control;
        FileSystemOptions options{};
        options.sizeBytes        = sizeof(options);
        options.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
        options.operationControl = &control;
        if (! DebugCheck(SUCCEEDED(readerControl->SetOperationControl(&options)), L"the reader should accept the operation control", passed, failed))
        {
            return;
        }

        std::atomic<HRESULT> readHr{E_PENDING};
        std::atomic<unsigned long> readBytes{0u};
        std::vector<std::byte> buffer(64u * 1024u);
        std::thread readerThread([&]() noexcept
        {
            unsigned long bytesRead = 0u;
            const HRESULT rh        = reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &bytesRead);
            readBytes.store(bytesRead, std::memory_order_release);
            readHr.store(rh, std::memory_order_release);
        });

        const ULONGLONG waitStart = GetTickCount64();
        while (CountCommands(endpoint.Snapshot(), "RETR") == 0u && GetTickCount64() - waitStart < 10'000u)
        {
            Sleep(20u);
        }
        const bool retrReached = CountCommands(endpoint.Snapshot(), "RETR") != 0u;
        DebugCheck(retrReached, L"the stalled RETR should be reached while Read waits", passed, failed);

        const ULONGLONG cancelTick = GetTickCount64();
        control.abortRequested.store(true, std::memory_order_release);
        while (readHr.load(std::memory_order_acquire) == E_PENDING && GetTickCount64() - cancelTick < 10'000u)
        {
            Sleep(20u);
        }
        const ULONGLONG cancelMs = GetTickCount64() - cancelTick;
        if (readHr.load(std::memory_order_acquire) == E_PENDING)
        {
            endpoint.Stop(); // release the wedged reply so the process can exit; the checks below record the failure
        }
        readerThread.join();

        DebugCheck(readHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                   L"a canceled stalled Read must report ERROR_CANCELLED through the reader's operation control",
                   passed,
                   failed);
        DebugCheck(cancelMs < 3'000u, L"the canceled stalled Read must return within the control's poll bound", passed, failed);
        Debug::Perf::EmitValue(L"FileSystemCurl.SelfTest.StalledReaderCancelMs", cancelMs, readHr.load(std::memory_order_acquire));
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported-selftest boundary: record the failure instead of escaping through C ABI.
        Debug::Error(L"FileSystemCurl stalled-reader selftest failed after std::exception.");
        DebugCheck(false, L"stalled-reader proof should not throw std::exception", passed, failed);
    }
}

namespace
{
class CancelControl final : public IFileSystemOperationControl
{
public:
    CancelControl() noexcept                       = default;
    CancelControl(const CancelControl&)            = delete;
    CancelControl& operator=(const CancelControl&) = delete;
    CancelControl(CancelControl&&)                 = delete;
    CancelControl& operator=(CancelControl&&)      = delete;

    HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void*) noexcept override
    {
        if (abort == nullptr)
        {
            return E_POINTER;
        }
        *abort = abortRequested.load(std::memory_order_acquire) ? TRUE : FALSE;
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
    std::atomic<bool> abortRequested{false};
};
} // namespace

void RunCurlStalledControlCommandCancelSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"WSAStartup should initialize the stalled-command proof", passed, failed))
    {
        return;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });

    try
    {
        FakeFtpEndpoint endpoint(UploadRetention::Complete);
        endpoint.SeedFile("/stall.bin", std::vector<uint8_t>(64u, 0x5Au));
        endpoint.SetStallVerb("DELE", 30'000u);
        if (! DebugCheck(SUCCEEDED(endpoint.Start()), L"stalled-command fixture should start", passed, failed))
        {
            return;
        }
        auto stopEndpoint = wil::scope_exit([&]() noexcept { endpoint.Stop(); });

        wil::com_ptr<IFileSystem> fileSystem;
        HRESULT hr = CreateFtpSelfTestInstance(fileSystem.put());
        if (! DebugCheck(SUCCEEDED(hr) && fileSystem, L"stalled-command fixture should create an FTP instance", passed, failed))
        {
            return;
        }
        wil::com_ptr<IInformations> information;
        hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
        if (! DebugCheck(SUCCEEDED(hr) && information, L"FTP instance should expose IInformations", passed, failed))
        {
            return;
        }
        // operationTimeoutMs 2000 -> low-speed/response bound of 2 s: the provider-owned watchdog.
        hr = information->SetConfiguration(
            R"({"connectTimeoutMs":2000,"operationTimeoutMs":2000,"copyMoveMaxConcurrency":1,"deleteMaxConcurrency":1,"ftpUseEpsv":true})");
        if (! DebugCheck(SUCCEEDED(hr), L"FTP instance should accept the stalled-command configuration", passed, failed))
        {
            return;
        }

        const std::wstring stalledPath = std::format(L"//anonymous@127.0.0.1:{}/stall.bin", endpoint.Port());
        const wchar_t* paths[]         = {stalledPath.c_str()};

        // 1) Cancel returns the wedged command in about a second.
        CancelControl control;
        FileSystemOptions options{};
        options.sizeBytes        = sizeof(options);
        options.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
        options.operationControl = &control;
        std::atomic<HRESULT> deleteHr{E_PENDING};
        std::thread deleter([&]() noexcept
        { deleteHr.store(fileSystem->DeleteItems(paths, 1u, FILESYSTEM_FLAG_NONE, &options, nullptr, nullptr), std::memory_order_release); });
        const ULONGLONG waitStart = GetTickCount64();
        while (CountCommands(endpoint.Snapshot(), "DELE") == 0u && GetTickCount64() - waitStart < 10'000u)
        {
            Sleep(20u);
        }
        const bool deleReached = CountCommands(endpoint.Snapshot(), "DELE") != 0u;
        Sleep(200u);
        const ULONGLONG cancelTick = GetTickCount64();
        control.abortRequested.store(true, std::memory_order_release);
        const bool returned      = WaitForSingleObject(deleter.native_handle(), 10'000u) == WAIT_OBJECT_0;
        const ULONGLONG cancelMs = GetTickCount64() - cancelTick;
        if (! returned)
        {
            endpoint.Stop(); // release the wedged reply so the process can exit; the checks below record the failure
        }
        deleter.join();
        DebugCheck(deleReached, L"the stalled DELE must reach the fixture", passed, failed);
        DebugCheck(returned, L"a control command the server never answers must return after Cancel", passed, failed);
        DebugCheck(deleteHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                   L"a canceled stalled control command must report ERROR_CANCELLED",
                   passed,
                   failed);
        DebugCheck(cancelMs < 3'000u, L"Cancel must return a stalled control command within a few progress-callback ticks", passed, failed);
        if (! returned)
        {
            return;
        }

        // 2) Without Cancel the provider-owned bound (response timeout) returns it.
        const ULONGLONG boundStart = GetTickCount64();
        const HRESULT boundHr      = fileSystem->DeleteItems(paths, 1u, FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        const ULONGLONG boundMs    = GetTickCount64() - boundStart;
        DebugCheck(FAILED(boundHr) && boundHr != HRESULT_FROM_WIN32(ERROR_CANCELLED),
                   L"an un-canceled stalled control command must fail through the transport bound, not as a cancel",
                   passed,
                   failed);
        // 2 s response bound per attempt; transient timeouts are retried up to three times with short backoff.
        DebugCheck(boundMs < 15'000u, L"the provider-owned bound must return a stalled control command on its own", passed, failed);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"FileSystemCurl stalled-command selftest failed after std::exception.");
        DebugCheck(false, L"stalled-command proof should not throw std::exception", passed, failed);
    }
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlNativeDeleteGuardsForSelfTest(unsigned int* passed, unsigned int* failed) noexcept
{
    if (passed == nullptr || failed == nullptr)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"native Delete fixture initializes Winsock", *passed, *failed))
    {
        return E_FAIL;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    try
    {
        const std::vector<uint8_t> sentinel{0x31u, 0xA7u, 0x00u, 0xFEu};
        for (unsigned int concurrency : {1u, 4u})
        {
            for (bool bulk : {false, true})
            {
                for (bool recursive : {false, true})
                {
                    FakeFtpEndpoint endpoint(UploadRetention::Complete);
                    endpoint.SeedFile("/sentinel.bin", sentinel);
                    endpoint.SeedFile("/protected/nested/child.bin", sentinel);
                    const EndpointSnapshot before = endpoint.Snapshot();
                    wil::com_ptr<IFileSystem> fileSystem;
                    HRESULT hr = endpoint.Start();
                    if (SUCCEEDED(hr))
                    {
                        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
                    }
                    wil::com_ptr<IInformations> information;
                    if (SUCCEEDED(hr))
                    {
                        hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
                    }
                    if (SUCCEEDED(hr))
                    {
                        hr = information->SetConfiguration(
                            std::format(
                                R"({{"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":1,"deleteMaxConcurrency":{},"ftpUseEpsv":true}})",
                                concurrency)
                                .c_str());
                    }
                    if (! DebugCheck(SUCCEEDED(hr), L"native root Delete fixture starts", *passed, *failed))
                    {
                        return E_FAIL;
                    }
                    const std::wstring root = std::format(L"//anonymous@127.0.0.1:{}/", endpoint.Port());
                    const wchar_t* paths[]  = {root.c_str()};
                    const auto flags        = recursive ? FILESYSTEM_FLAG_RECURSIVE : FILESYSTEM_FLAG_NONE;
                    const auto started      = std::chrono::steady_clock::now();
                    hr                      = bulk ? fileSystem->DeleteItems(paths, 1u, flags, nullptr, nullptr, nullptr)
                                                   : fileSystem->DeleteItem(root.c_str(), flags, nullptr, nullptr, nullptr);
                    const uint64_t durationUs =
                        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());
                    const EndpointSnapshot after = endpoint.Snapshot();
                    const size_t mutations       = CountCommands(after, "DELE") + CountCommands(after, "RMD");
                    const bool guarded           = hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) && mutations == 0u && before.files == after.files &&
                                                   before.directories == after.directories;
                    const std::wstring route = std::format(L"{};concurrency={};recursive={}", bulk ? L"DeleteItems" : L"DeleteItem", concurrency, recursive);
                    Debug::Perf::Emit(
                        L"FileOps.Curl.NativeDelete.RootGuard", route.c_str(), durationUs, mutations, after.files.size(), guarded ? S_OK : E_FAIL);
                    DebugCheck(guarded, std::format(L"{} must refuse root before any mutation and preserve every descendant", route).c_str(), *passed, *failed);

                    // A root guard must not disable ordinary native Delete. Remove the selected
                    // subtree through the same entry point and keep its neighboring file intact.
                    const std::wstring selected    = std::format(L"//anonymous@127.0.0.1:{}/protected", endpoint.Port());
                    const wchar_t* selectedPaths[] = {selected.c_str()};
                    const auto deleteStarted       = std::chrono::steady_clock::now();
                    const HRESULT selectedHr       = bulk ? fileSystem->DeleteItems(selectedPaths, 1u, FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr)
                                                          : fileSystem->DeleteItem(selected.c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr);
                    const uint64_t selectedUs =
                        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - deleteStarted).count());
                    const EndpointSnapshot selectedAfter = endpoint.Snapshot();
                    const bool contained = selectedHr == S_OK && selectedAfter.files.size() == 1u && FileEquals(selectedAfter, "/sentinel.bin", sentinel) &&
                                           selectedAfter.directories == std::set<std::string>{"/"} && CountCommands(selectedAfter, "DELE") == 1u &&
                                           CountCommands(selectedAfter, "RMD") == 2u;
                    Debug::Perf::Emit(L"FileOps.Curl.NativeDelete.SelectedSubtree",
                                      route.c_str(),
                                      selectedUs,
                                      CountCommands(selectedAfter, "DELE"),
                                      CountCommands(selectedAfter, "RMD"),
                                      contained ? S_OK : E_FAIL);
                    DebugCheck(contained, L"native Delete removes only the selected subtree and preserves its sibling", *passed, *failed);
                }
            }
        }

        for (bool bulk : {false, true})
        {
            FakeFtpEndpoint endpoint(UploadRetention::Complete);
            endpoint.SeedFile("/selected/child.bin", sentinel);
            endpoint.SeedFile("/sibling.bin", sentinel);
            endpoint.SetStallVerb("LIST", 30'000u);
            const EndpointSnapshot before = endpoint.Snapshot();
            wil::com_ptr<IFileSystem> fileSystem;
            HRESULT hr = endpoint.Start();
            if (SUCCEEDED(hr))
            {
                hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
            }
            if (! DebugCheck(SUCCEEDED(hr), L"native Delete listing-cancel fixture starts", *passed, *failed))
            {
                return E_FAIL;
            }
            const std::wstring selected = std::format(L"//anonymous@127.0.0.1:{}/selected", endpoint.Port());
            const wchar_t* paths[]      = {selected.c_str()};
            CancelControl control;
            FileSystemOptions options{};
            options.sizeBytes        = sizeof(options);
            options.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
            options.operationControl = &control;
            HRESULT deleteHr         = E_PENDING;
            std::jthread worker([&]() noexcept
            {
                deleteHr = bulk ? fileSystem->DeleteItems(paths, 1u, FILESYSTEM_FLAG_RECURSIVE, &options, nullptr, nullptr)
                                : fileSystem->DeleteItem(selected.c_str(), FILESYSTEM_FLAG_RECURSIVE, &options, nullptr, nullptr);
            });
            const ULONGLONG deadline = GetTickCount64() + 5'000u;
            while (CountCommands(endpoint.Snapshot(), "LIST") == 0u && GetTickCount64() < deadline)
            {
                Sleep(10u);
            }
            const bool reached         = CountCommands(endpoint.Snapshot(), "LIST") != 0u;
            const ULONGLONG cancelTick = GetTickCount64();
            control.abortRequested.store(true, std::memory_order_release);
            const bool returned     = WaitForSingleObject(worker.native_handle(), 2'500u) == WAIT_OBJECT_0;
            const uint64_t cancelMs = GetTickCount64() - cancelTick;
            if (! returned)
            {
                endpoint.Stop(); // Fixture-owned unblock only; never counts as cooperative cancellation.
            }
            worker.join();
            const EndpointSnapshot after = endpoint.Snapshot();
            const size_t mutations       = CountCommands(after, "DELE") + CountCommands(after, "RMD");
            const bool canceled          = reached && returned && cancelMs < 2'000u && deleteHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && mutations == 0u &&
                                           before.files == after.files && before.directories == after.directories;
            Debug::Perf::Emit(L"FileOps.Curl.NativeDelete.ListCancel",
                              bulk ? L"DeleteItems" : L"DeleteItem",
                              cancelMs * 1'000u,
                              mutations,
                              returned ? 1u : 0u,
                              canceled ? S_OK : E_FAIL);
            DebugCheck(canceled,
                       bulk ? L"DeleteItems cancels its stalled admission listing without mutation"
                            : L"DeleteItem cancels its stalled admission listing without mutation",
                       *passed,
                       *failed);
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: record failure without escaping into the host.
        Debug::Error(L"Curl native Delete guard fixture failed after std::exception.");
        DebugCheck(false, L"native Delete guard fixture must not throw", *passed, *failed);
    }
    return *failed == 0u ? S_OK : E_FAIL;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlNativeDeleteLateListingForSelfTest(unsigned int* passed, unsigned int* failed) noexcept
{
    if (passed == nullptr || failed == nullptr)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"late-listing fixture initializes Winsock", *passed, *failed))
    {
        return E_FAIL;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    try
    {
        const std::vector<uint8_t> bytes{0x31u, 0x00u, 0xAAu, 0xFEu};
        for (const LateListingReply reply : {LateListingReply::Complete, LateListingReply::Refuse, LateListingReply::WaitForCancel})
        {
            for (const bool bulk : {false, true})
            {
                for (const unsigned int concurrency : {1u, 4u})
                {
                    FakeFtpEndpoint endpoint(UploadRetention::Complete);
                    endpoint.SeedFile("/selected/a/removed.bin", bytes);
                    endpoint.SeedFile("/selected/b/retained.bin", bytes);
                    endpoint.SeedFile("/sibling.bin", bytes);
                    endpoint.SetLateListingGate("/selected/b", "/selected/a/removed.bin", reply);
                    wil::com_ptr<IFileSystem> fileSystem;
                    HRESULT hr = endpoint.Start();
                    if (SUCCEEDED(hr))
                    {
                        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
                    }
                    wil::com_ptr<IInformations> information;
                    if (SUCCEEDED(hr))
                    {
                        hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
                    }
                    if (SUCCEEDED(hr))
                    {
                        hr = information->SetConfiguration(
                            std::format(
                                R"({{"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":1,"deleteMaxConcurrency":{},"ftpUseEpsv":true}})",
                                concurrency)
                                .c_str());
                    }
                    FileSystemArenaOwner arena;
                    if (SUCCEEDED(hr))
                    {
                        hr = arena.Initialize(4096u);
                    }
                    if (! DebugCheck(SUCCEEDED(hr), L"late-listing native Delete fixture starts", *passed, *failed))
                    {
                        return E_FAIL;
                    }
                    const wchar_t* source = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/selected", endpoint.Port()));
                    auto* paths = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena.Get(), sizeof(const wchar_t*), alignof(const wchar_t*)));
                    if (source == nullptr || paths == nullptr)
                    {
                        DebugCheck(false, L"late-listing fixture fits its bounded caller arena", *passed, *failed);
                        return E_FAIL;
                    }
                    paths[0] = source;
                    CleanupDebtOperationCallback callback;
                    CancelControl control;
                    FileSystemOptions options{};
                    options.sizeBytes        = sizeof(options);
                    options.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
                    options.operationControl = &control;
                    HRESULT deleteHr         = E_PENDING;
                    const auto started       = std::chrono::steady_clock::now();
                    std::jthread worker([&]() noexcept
                    {
                        deleteHr = bulk ? fileSystem->DeleteItems(paths, 1u, FILESYSTEM_FLAG_RECURSIVE, &options, &callback, nullptr)
                                        : fileSystem->DeleteItem(source, FILESYSTEM_FLAG_RECURSIVE, &options, &callback, nullptr);
                    });
                    bool canceledCooperatively = false;
                    uint64_t cancelUs          = 0u;
                    if (reply == LateListingReply::WaitForCancel)
                    {
                        const bool reachedAfterDelete = endpoint.WaitForLateListingRelease(6s);
                        const auto cancelStarted      = std::chrono::steady_clock::now();
                        control.abortRequested.store(true, std::memory_order_release);
                        const bool returned   = WaitForSingleObject(worker.native_handle(), 2'500u) == WAIT_OBJECT_0;
                        cancelUs              = Debug::Perf::ElapsedUs(cancelStarted);
                        canceledCooperatively = reachedAfterDelete && returned && cancelUs < 2'000'000u;
                        if (! returned)
                        {
                            endpoint.Stop(); // Owned emergency unblock is never cooperative-cancel evidence.
                        }
                    }
                    worker.join();
                    const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
                    endpoint.Stop();
                    const EndpointSnapshot after = endpoint.Snapshot();
                    const bool ordered   = after.lateListingReachedCount == 1u && after.lateListingReleasedCount == 1u && after.lateListingTimeoutCount == 0u;
                    const bool completed = reply == LateListingReply::Complete;
                    const bool failureState  = after.files.size() == 2u && FileAbsent(after, "/selected/a/removed.bin") &&
                                               FileEquals(after, "/selected/b/retained.bin", bytes) && after.directories.contains("/selected") &&
                                               after.deletedPaths == std::vector<std::string>{"/selected/a/removed.bin"};
                    const bool successState  = after.files.size() == 1u && ! after.directories.contains("/selected") && after.deletedPaths.size() == 2u;
                    const bool resultCorrect = completed ? deleteHr == S_OK
                                               : reply == LateListingReply::WaitForCancel
                                                   ? canceledCooperatively && deleteHr == HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                   : FAILED(deleteHr) && deleteHr != HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    const bool stateCorrect =
                        resultCorrect && SUCCEEDED(after.serverHr) && FileEquals(after, "/sibling.bin", bytes) && (completed ? successState : failureState);
                    const bool truthful      = callback.SingleCompletionHasTruth(! completed);
                    const std::wstring route = std::format(L"{};concurrency={};reply={}",
                                                           bulk ? L"DeleteItems" : L"DeleteItem",
                                                           concurrency,
                                                           completed                           ? L"complete"
                                                           : reply == LateListingReply::Refuse ? L"refuse"
                                                                                               : L"cancel");
                    Debug::Perf::Emit(L"FileOps.Curl.NativeDelete.LateListing",
                                      route.c_str(),
                                      durationUs,
                                      after.lateListingReleasedCount,
                                      after.lateListingTimeoutCount,
                                      ordered && stateCorrect && truthful ? S_OK : E_FAIL);
                    if (reply == LateListingReply::WaitForCancel)
                    {
                        Debug::Perf::Emit(L"FileOps.Curl.NativeDelete.LateListCancel",
                                          route.c_str(),
                                          cancelUs,
                                          after.deletedPaths.size(),
                                          canceledCooperatively ? 1u : 0u,
                                          stateCorrect ? S_OK : E_FAIL);
                    }
                    DebugCheck(ordered, std::format(L"{} deletes known work before later discovery completes", route).c_str(), *passed, *failed);
                    DebugCheck(stateCorrect,
                               std::format(L"{} preserves siblings and reports the exact completion/cancel/failure state", route).c_str(),
                               *passed,
                               *failed);
                    DebugCheck(truthful, std::format(L"{} publishes one truthful terminal mutation receipt", route).c_str(), *passed, *failed);
                }
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: make fixture failure visible to the host.
        Debug::Error(L"Curl native Delete late-listing fixture failed after std::exception.");
        DebugCheck(false, L"late-listing native Delete fixture must not throw", *passed, *failed);
    }
    return *failed == 0u ? S_OK : E_FAIL;
}

namespace
{
// Fixture-specific raw-vtable observer: unlike an operation callback or a UI
// message-pump probe, this verifies size totals, cookie, and progress/cancel order.
// Only cancelRequested crosses threads; all other observations are read after join.
class CurlSizeTruthCallback final : public IFileSystemDirectorySizeCallback
{
public:
    CurlSizeTruthCallback() noexcept                               = default;
    CurlSizeTruthCallback(const CurlSizeTruthCallback&)            = delete;
    CurlSizeTruthCallback& operator=(const CurlSizeTruthCallback&) = delete;
    CurlSizeTruthCallback(CurlSizeTruthCallback&&)                 = delete;
    CurlSizeTruthCallback& operator=(CurlSizeTruthCallback&&)      = delete;

    HRESULT STDMETHODCALLTYPE
    DirectorySizeProgress(uint64_t scanned, uint64_t bytes, uint64_t files, uint64_t directories, const wchar_t* currentPath, void* cookie) noexcept override
    {
        contractCorrect = contractCorrect && cookie == this && finalCount == 0u && scanned >= lastScanned && bytes >= lastBytes && files >= lastFiles &&
                          directories >= lastDirectories;
        lastScanned     = scanned;
        lastBytes       = bytes;
        lastFiles       = files;
        lastDirectories = directories;
        ++progressCount;
        progressPending = true;
        if (scanned >= cancelAfterScanned)
        {
            cancelRequested.store(true, std::memory_order_release);
        }
        if (! currentPath)
        {
            ++finalCount;
            if (cancelOnFinal)
            {
                cancelRequested.store(true, std::memory_order_release);
            }
            return finalProgressResult;
        }
        if (scanned >= failAfterScanned)
        {
            return E_ACCESSDENIED;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DirectorySizeShouldCancel(BOOL* cancel, void* cookie) noexcept override
    {
        if (! cancel)
        {
            return E_POINTER;
        }
        contractCorrect = contractCorrect && cookie == this && progressPending;
        progressPending = false;
        ++cancelChecks;
        *cancel = cancelRequested.load(std::memory_order_acquire) ? TRUE : FALSE;
        return S_OK;
    }

    std::atomic<bool> cancelRequested{false};
    uint64_t cancelAfterScanned = (std::numeric_limits<uint64_t>::max)();
    uint64_t failAfterScanned   = (std::numeric_limits<uint64_t>::max)();
    bool cancelOnFinal          = false;
    HRESULT finalProgressResult = S_OK;
    uint64_t lastScanned        = 0u;
    uint64_t lastBytes          = 0u;
    uint64_t lastFiles          = 0u;
    uint64_t lastDirectories    = 0u;
    uint64_t progressCount      = 0u;
    uint64_t cancelChecks       = 0u;
    unsigned int finalCount     = 0u;
    bool progressPending        = false;
    bool contractCorrect        = true;
};
} // namespace

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlDirectorySizeTruthForSelfTest(unsigned int* passed, unsigned int* failed) noexcept
{
    if (! passed || ! failed)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"directory-size fixture initializes Winsock", *passed, *failed))
    {
        return E_FAIL;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    try
    {
        enum class Shape
        {
            FileKnown,
            FileProbe,
            FileUnknown,
            ChildProbe,
            ChildUnknown,
            Empty,
            Shallow,
            Wide,
            Deep,
            DepthLimit,
            PreCanceled,
            FlatCancel,
            RootFailurePrecedence,
            ChildDenied,
            Malformed,
            Outside,
            Overlong,
            ByteOverflow,
            StalledListCancel,
            LiteralArrow,
            CallbackFailure,
            FinalCancel,
            PartialFinalCancel,
            PartialFinalFailure,
            LateMalformed,
            WideParent,
        };
        const std::vector<uint8_t> bytes{0x31u, 0x00u, 0xFFu};
        for (const Shape shape : {Shape::FileKnown,
                                  Shape::FileProbe,
                                  Shape::FileUnknown,
                                  Shape::ChildProbe,
                                  Shape::ChildUnknown,
                                  Shape::Empty,
                                  Shape::Shallow,
                                  Shape::Wide,
                                  Shape::Deep,
                                  Shape::DepthLimit,
                                  Shape::PreCanceled,
                                  Shape::FlatCancel,
                                  Shape::RootFailurePrecedence,
                                  Shape::ChildDenied,
                                  Shape::Malformed,
                                  Shape::Outside,
                                  Shape::Overlong,
                                  Shape::ByteOverflow,
                                  Shape::StalledListCancel,
                                  Shape::LiteralArrow,
                                  Shape::CallbackFailure,
                                  Shape::FinalCancel,
                                  Shape::PartialFinalCancel,
                                  Shape::PartialFinalFailure,
                                  Shape::LateMalformed,
                                  Shape::WideParent})
        {
            const bool fileRoot          = shape == Shape::FileKnown || shape == Shape::FileProbe || shape == Shape::FileUnknown || shape == Shape::FinalCancel;
            const bool partialFinal      = shape == Shape::PartialFinalCancel || shape == Shape::PartialFinalFailure;
            const bool unknown           = shape == Shape::FileUnknown || shape == Shape::ChildUnknown || partialFinal;
            const bool missingListedSize = unknown || shape == Shape::FileProbe || shape == Shape::ChildProbe;
            FakeFtpEndpoint endpoint(UploadRetention::Complete, ! unknown, ! missingListedSize);
            const wchar_t* name          = L"known-file";
            HRESULT expectedHr           = S_OK;
            uint64_t expectedBytes       = bytes.size();
            uint64_t expectedFiles       = 1u;
            uint64_t expectedDirectories = 0u;
            FileSystemFlags flags        = FILESYSTEM_FLAG_RECURSIVE;
            CurlSizeTruthCallback callback;
            endpoint.SeedDirectory("/selected");
            if (fileRoot)
            {
                endpoint.SeedFile("/file.bin", bytes);
                name = shape == Shape::FileProbe ? L"file-size-probe" : shape == Shape::FileUnknown ? L"file-size-unknown" : L"known-file";
                if (unknown)
                {
                    expectedHr    = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                    expectedBytes = 0u;
                }
                if (shape == Shape::FinalCancel)
                {
                    name                   = L"cancel-from-final-progress";
                    callback.cancelOnFinal = true;
                    expectedHr             = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
            }
            else if (shape == Shape::Empty)
            {
                name          = L"empty-directory";
                expectedBytes = 0u;
                expectedFiles = 0u;
            }
            else if (shape == Shape::Wide || shape == Shape::FlatCancel)
            {
                name          = shape == Shape::Wide ? L"wide-4097" : L"flat-progress-cancel";
                expectedFiles = Common::FileOperations::kTraversalMaxQueuedEntries + 1u;
                expectedBytes = bytes.size() * expectedFiles;
                for (size_t index = 0u; index < expectedFiles; ++index)
                {
                    endpoint.SeedFile(std::format("/selected/leaf-{:05}.bin", index), bytes);
                }
                if (shape == Shape::FlatCancel)
                {
                    callback.cancelAfterScanned = 17u;
                    expectedHr                  = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
            }
            else if (shape == Shape::Deep || shape == Shape::DepthLimit)
            {
                name                  = shape == Shape::Deep ? L"deep-80" : L"depth-limit-129";
                expectedDirectories   = shape == Shape::Deep ? 80u : Common::FileOperations::kTraversalMaxDepth + 1u;
                std::string directory = "/selected";
                for (size_t index = 0u; index < expectedDirectories; ++index)
                {
                    directory.append("/d");
                }
                endpoint.SeedFile(directory + "/leaf.bin", bytes);
                if (shape == Shape::DepthLimit)
                {
                    expectedHr    = HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
                    expectedBytes = 0u;
                    expectedFiles = 0u;
                }
            }
            else if (shape == Shape::CallbackFailure)
            {
                name = L"callback-error-is-global";
                for (size_t index = 0u; index < 1000u; ++index)
                {
                    endpoint.SeedFile(std::format("/selected/a/leaf-{:05}.bin", index), bytes);
                }
                endpoint.SeedFile("/selected/z/leaf.bin", bytes);
                callback.failAfterScanned = 100u;
                expectedHr                = E_ACCESSDENIED;
            }
            else if (shape == Shape::ChildDenied)
            {
                name = L"denied-child-continues-siblings";
                endpoint.SeedFile("/selected/a/leaf.bin", bytes);
                endpoint.SeedFile("/selected/b/leaf.bin", bytes);
                endpoint.SeedFile("/selected/z/leaf.bin", bytes);
                endpoint.SetDirectoryAccessDenied("/selected/b");
                expectedHr          = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                expectedBytes       = 6u;
                expectedFiles       = 2u;
                expectedDirectories = 3u;
            }
            else
            {
                endpoint.SeedFile(shape == Shape::LiteralArrow ? "/selected/ordinary -> suffix.bin" : "/selected/leaf.bin", bytes);
                if (shape == Shape::ChildProbe || shape == Shape::ChildUnknown || partialFinal)
                {
                    name = unknown ? L"child-size-unknown" : L"child-size-probe";
                    if (unknown)
                    {
                        expectedHr    = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                        expectedBytes = 0u;
                    }
                    if (partialFinal)
                    {
                        const bool cancel            = shape == Shape::PartialFinalCancel;
                        name                         = cancel ? L"partial-total-final-cancel" : L"partial-total-final-callback-error";
                        callback.cancelOnFinal       = cancel;
                        callback.finalProgressResult = cancel ? S_OK : E_ACCESSDENIED;
                        expectedHr                   = cancel ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : E_ACCESSDENIED;
                    }
                }
                else if (shape == Shape::Shallow)
                {
                    name = L"immediate-children-only";
                    endpoint.SeedFile("/selected/nested/leaf.bin", bytes);
                    expectedDirectories = 1u;
                    flags               = FILESYSTEM_FLAG_NONE;
                }
                else if (shape == Shape::ByteOverflow)
                {
                    name = L"checked-byte-overflow";
                    endpoint.SetDirectoryPayload("/selected",
                                                 "-rw-r--r-- 1 owner group 18446744073709551615 Jan 01 2026 a.bin\r\n"
                                                 "-rw-r--r-- 1 owner group 1 Jan 01 2026 b.bin\r\n");
                    expectedHr    = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                    expectedBytes = (std::numeric_limits<uint64_t>::max)();
                    expectedFiles = 2u;
                }
                else if (shape == Shape::LiteralArrow)
                {
                    name = L"literal-arrow-name";
                }
                else if (shape == Shape::LateMalformed)
                {
                    name = L"late-malformed-list-retains-known-total";
                    endpoint.SetDirectoryPayload("/selected", "-rw-r--r-- 1 owner group 3 Jan 01 2026 leaf.bin\r\nnot a supported LIST record\r\n");
                    expectedHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                }
                else if (shape == Shape::WideParent)
                {
                    name = L"wide-parent-root-lookup";
                    for (size_t index = 0u; index < Common::FileOperations::kTraversalMaxQueuedEntries + 1u; ++index)
                    {
                        endpoint.SeedFile(std::format("/unselected-{:05}.bin", index), bytes);
                    }
                }
                else
                {
                    expectedBytes = 0u;
                    expectedFiles = 0u;
                    if (shape == Shape::PreCanceled)
                    {
                        name = L"pre-cancel-before-network";
                        callback.cancelRequested.store(true, std::memory_order_release);
                        expectedHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }
                    else if (shape == Shape::StalledListCancel)
                    {
                        name = L"cancel-during-list-wait";
                        endpoint.SetLateListingGate("/selected", "/never-deleted", LateListingReply::WaitForCancel);
                        expectedHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }
                    else if (shape == Shape::RootFailurePrecedence)
                    {
                        name = L"root-error-before-final-callback-error";
                        endpoint.SetDirectoryAccessDenied("/selected");
                        callback.finalProgressResult = E_ABORT;
                        expectedHr                   = E_ACCESSDENIED;
                    }
                    else if (shape == Shape::Malformed)
                    {
                        name = L"malformed-list";
                        endpoint.SetDirectoryPayload("/selected", "not a supported LIST record\r\n");
                        expectedHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    }
                    else if (shape == Shape::Outside)
                    {
                        name = L"invalid-child-path";
                        endpoint.SetDirectoryPayload("/selected", "-rw-r--r-- 1 owner group 3 Jan 01 2026 ../sibling.bin\r\n");
                        expectedHr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                    }
                    else if (shape == Shape::Overlong)
                    {
                        name = L"overlong-list-row";
                        endpoint.SetDirectoryPayload("/selected", std::string(CURL_MAX_WRITE_SIZE + 1u, 'x') + "\r\n");
                        expectedHr = HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
                    }
                }
            }
            endpoint.SeedFile("/sibling.bin", bytes);
            const EndpointSnapshot before = endpoint.Snapshot();
            wil::com_ptr<IFileSystem> fileSystem;
            wil::com_ptr<IFileSystemDirectoryOperations> directoryOperations;
            HRESULT hr = endpoint.Start();
            if (SUCCEEDED(hr))
            {
                hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
            }
            if (SUCCEEDED(hr))
            {
                hr = fileSystem->QueryInterface(IID_PPV_ARGS(directoryOperations.put()));
            }
            if (! DebugCheck(SUCCEEDED(hr), std::format(L"{} fixture starts", name).c_str(), *passed, *failed))
            {
                return E_FAIL;
            }
            const std::wstring path = std::format(L"//anonymous@127.0.0.1:{}/{}", endpoint.Port(), fileRoot ? L"file.bin" : L"selected");
            if (shape == Shape::FileKnown)
            {
                FileSystemDirectorySizeResult invalid{};
                invalid.sizeBytes = sizeof(invalid) - 1u;
                DebugCheck(directoryOperations->GetDirectorySize(path.c_str(), flags, nullptr, nullptr, nullptr) == E_POINTER,
                           L"size rejects a null result",
                           *passed,
                           *failed);
                DebugCheck(directoryOperations->GetDirectorySize(path.c_str(), flags, nullptr, nullptr, &invalid) == E_INVALIDARG,
                           L"size rejects an incompatible result prefix",
                           *passed,
                           *failed);
                invalid.sizeBytes = sizeof(invalid);
                DebugCheck(directoryOperations->GetDirectorySize(nullptr, flags, nullptr, nullptr, &invalid) == E_POINTER && invalid.status == E_POINTER,
                           L"size rejects a null root consistently",
                           *passed,
                           *failed);
                DebugCheck(directoryOperations->GetDirectorySize(L"", flags, nullptr, nullptr, &invalid) == E_INVALIDARG && invalid.status == E_INVALIDARG &&
                               endpoint.Snapshot().commands.empty(),
                           L"size rejects an empty root without network I/O",
                           *passed,
                           *failed);
            }
            FileSystemDirectorySizeResult result{};
            result.sizeBytes       = sizeof(result);
            uint64_t cancelUs      = 0u;
            bool cooperativeCancel = true;
            const auto started     = std::chrono::steady_clock::now();
            if (shape == Shape::StalledListCancel)
            {
                std::jthread worker([&]() noexcept { hr = directoryOperations->GetDirectorySize(path.c_str(), flags, &callback, &callback, &result); });
                const bool reached       = endpoint.WaitForLateListingReached(3s);
                const auto cancelStarted = std::chrono::steady_clock::now();
                callback.cancelRequested.store(true, std::memory_order_release);
                const bool returned = WaitForSingleObject(worker.native_handle(), 2'000u) == WAIT_OBJECT_0;
                cancelUs            = Debug::Perf::ElapsedUs(cancelStarted);
                cooperativeCancel   = reached && returned && cancelUs < 2'000'000u;
                if (! returned)
                {
                    endpoint.Stop(); // Owned emergency unblock is not cooperative-cancel evidence.
                }
                worker.join();
            }
            else
            {
                hr = directoryOperations->GetDirectorySize(path.c_str(), flags, &callback, &callback, &result);
            }
            const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
            endpoint.Stop();
            const EndpointSnapshot after = endpoint.Snapshot();
            const bool totalsCorrect =
                shape == Shape::CallbackFailure
                    // Parent LIST finishes before child Open, so sibling `z` is already
                    // observed when the callback aborts inside `a`. `z`'s files stay
                    // unsized; the abort remains global.
                    ? result.fileCount >= 90u && result.fileCount <= 199u && result.totalBytes == result.fileCount * bytes.size() && result.directoryCount == 2u
                : shape == Shape::FlatCancel
                    ? result.fileCount >= 17u && result.fileCount <= 128u && result.totalBytes == result.fileCount * bytes.size() && result.directoryCount == 0u
                    : result.totalBytes == expectedBytes && result.fileCount == expectedFiles && result.directoryCount == expectedDirectories;
            const bool contentsCorrect = before.files == after.files && before.directories == after.directories && SUCCEEDED(after.serverHr);
            const bool correct         = hr == expectedHr && result.status == hr && totalsCorrect && contentsCorrect && cooperativeCancel &&
                                         (shape != Shape::PreCanceled || after.commands.empty());
            const bool feedback        = callback.contractCorrect && callback.progressCount != 0u && callback.cancelChecks != 0u && callback.finalCount <= 1u &&
                                         callback.lastBytes == result.totalBytes && callback.lastFiles == result.fileCount &&
                                         callback.lastDirectories == result.directoryCount &&
                                         (hr != S_OK || (callback.finalCount == 1u && ! callback.progressPending));
            const std::wstring detail  = std::format(L"{};hr=0x{:08X};status=0x{:08X};expected=0x{:08X};server=0x{:08X};bytes={};files={};dirs={};progress={};"
                                                     L"checks={};final={};order={};state={};feedback={}",
                                                     name,
                                                     static_cast<unsigned long>(hr),
                                                     static_cast<unsigned long>(result.status),
                                                     static_cast<unsigned long>(expectedHr),
                                                     static_cast<unsigned long>(after.serverHr),
                                                     result.totalBytes,
                                                     result.fileCount,
                                                     result.directoryCount,
                                                     callback.progressCount,
                                                     callback.cancelChecks,
                                                     callback.finalCount,
                                                     callback.contractCorrect,
                                                     correct,
                                                     feedback);
            Debug::Perf::Emit(L"FileOps.Curl.DirectorySize.TruthFixture",
                              detail.c_str(),
                              durationUs,
                              result.fileCount,
                              result.directoryCount,
                              correct && feedback ? S_OK : E_FAIL);
            if (shape == Shape::StalledListCancel)
            {
                Debug::Perf::Emit(L"FileOps.Curl.DirectorySize.ListCancel",
                                  name,
                                  cancelUs,
                                  after.lateListingReachedCount,
                                  cooperativeCancel ? 1u : 0u,
                                  correct ? S_OK : E_FAIL);
            }
            DebugCheck(correct, std::format(L"{} preserves contents and exact size/partial/error/cancel truth", name).c_str(), *passed, *failed);
            DebugCheck(feedback, std::format(L"{} preserves ordered progress, cookie, and latest totals", name).c_str(), *passed, *failed);
            if (shape == Shape::FileKnown)
            {
                // The successful request initialized Curl. This private pool has
                // no other borrower or eviction; idleView is non-owning and stays
                // valid only until BeginShutdown below. Inspect before re-borrow
                // so a reset in Borrow cannot hide dangling per-call options.
                CurlEasyPool pool;
                CURL* idleView = nullptr;
                {
                    auto borrowed         = pool.Borrow(L"directory-size-idle-reset");
                    const bool configured = borrowed && curl_easy_setopt(borrowed.get(), CURLOPT_PRIVATE, &callback) == CURLE_OK;
                    if (! DebugCheck(configured, L"idle reset fixture stores per-call state", *passed, *failed))
                    {
                        return E_FAIL;
                    }
                    idleView = borrowed.get();
                }
                char* retained   = nullptr;
                const bool reset = curl_easy_getinfo(idleView, CURLINFO_PRIVATE, &retained) == CURLE_OK && retained == nullptr;
                DebugCheck(reset, L"pool clears per-call options before idle publication", *passed, *failed);
                Debug::Perf::Emit(
                    L"FileOps.Curl.DirectorySize.IdleHandleReset", L"options-cleared-before-reborrow", 0u, reset ? 1u : 0u, 0u, reset ? S_OK : E_FAIL);
                pool.BeginShutdown();
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: report fixture failure without unwinding into the host.
        Debug::Error(L"Curl directory-size truth fixture failed after std::exception.");
        DebugCheck(false, L"directory-size fixture must not throw", *passed, *failed);
    }
    return *failed == 0u ? S_OK : E_FAIL;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlMovePreflightTruthForSelfTest(unsigned int* passed, unsigned int* failed) noexcept
{
    if (! passed || ! failed)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"Move preflight fixture initializes Winsock", *passed, *failed))
    {
        return E_FAIL;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    {
        // Plan C0 item 1: the preflight's whole-operation memory is the digest containers, not one
        // string per member. 4,096 files under 64 directories with long names retain 16 bytes per
        // file and 8 per directory (plus vector headroom), answer membership and size proof exactly,
        // and refuse to answer before Finalize so a walker can never consult a half-built set.
        SourceTreeDigestSet members;
        SourceSizeCommitmentDigestMap commitments;
        const std::wstring root = NormalizePluginPath(L"/move-preflight-digest/very-long-directory-name-to-make-strings-expensive");
        members.Insert(root);
        std::vector<std::wstring> directories;
        for (unsigned int directory = 0u; directory < 64u; ++directory)
        {
            directories.push_back(JoinPluginPath(root, std::format(L"subdirectory-with-a-long-name-{:03}", directory)));
            members.Insert(directories.back());
            for (unsigned int file = 0u; file < 64u; ++file)
            {
                const std::wstring path = JoinPluginPath(directories.back(), std::format(L"file-with-a-long-name-{:03}-{:03}.bin", directory, file));
                members.Insert(path);
                commitments.Insert(path, static_cast<uint64_t>(directory) * 1000u + file);
            }
        }
        DebugCheck(! members.Contains(root) && ! commitments.Find(JoinPluginPath(directories[3], L"file-with-a-long-name-003-007.bin")).has_value(),
                   L"digest containers answer nothing before Finalize",
                   *passed,
                   *failed);
        commitments.Insert(JoinPluginPath(directories[3], L"file-with-a-long-name-003-007.bin"), 424242u);
        members.Finalize();
        commitments.Finalize();
        DebugCheck(members.IsFinalized() && commitments.IsFinalized() && members.size() == 1u + 64u + 4096u && commitments.size() == 4096u,
                   L"digest containers hold one entry per member after Finalize (duplicates collapse)",
                   *passed,
                   *failed);
        DebugCheck(members.Contains(root) && members.Contains(directories[63]) &&
                       members.Contains(JoinPluginPath(directories[63], L"file-with-a-long-name-063-063.bin")) &&
                       ! members.Contains(JoinPluginPath(directories[63], L"late-writer.bin")) && ! members.Contains(JoinPluginPath(root, L"late-dir")),
                   L"digest membership answers exactly for members and late writers",
                   *passed,
                   *failed);
        DebugCheck(commitments.Find(JoinPluginPath(directories[3], L"file-with-a-long-name-003-007.bin")) == std::optional<uint64_t>{424242u} &&
                       commitments.Find(JoinPluginPath(directories[9], L"file-with-a-long-name-009-011.bin")) == std::optional<uint64_t>{9011u} &&
                       ! commitments.Find(JoinPluginPath(directories[9], L"late-writer.bin")).has_value(),
                   L"digest size proof returns the last committed size and nothing for late writers",
                   *passed,
                   *failed);
        const uint64_t retained = members.RetainedBytes() + commitments.RetainedBytes();
        const uint64_t exact    = static_cast<uint64_t>(1u + 64u + 4096u) * sizeof(uint64_t) + 4096ull * sizeof(std::pair<uint64_t, uint64_t>);
        DebugCheck(retained >= exact && retained <= exact * 2u + 4096u,
                   std::format(L"digest containers retain about 16 bytes per file (retained={} exact={})", retained, exact).c_str(),
                   *passed,
                   *failed);
        Debug::Perf::Emit(L"FileOps.Curl.MovePreflight.RetainedDigestBytes", L"selftest-4096-files", 0u, retained, members.size() + commitments.size(), S_OK);
        DebugCheck(members.Contains(JoinPluginPath(NormalizePluginPath(L"/move-preflight-digest/very-long-directory-name-to-make-strings-expensive/"),
                                                   L"subdirectory-with-a-long-name-010")),
                   L"digest membership matches the walkers' normalize-then-join spelling",
                   *passed,
                   *failed);
    }
    try
    {
        enum class Shape
        {
            Shallow,
            Deep,
            DepthLimit,
            Malformed,
            Outside,
            Overlong,
            WideUnknown,
            StalledListCancel
        };
        const std::vector<uint8_t> bytes{0x31u, 0x00u, 0xAAu, 0xFEu};
        for (const Shape shape :
             {Shape::Shallow, Shape::Deep, Shape::Malformed, Shape::Outside, Shape::Overlong, Shape::WideUnknown, Shape::StalledListCancel, Shape::DepthLimit})
        {
            for (const bool bulk : {false, true})
            {
                for (const unsigned int concurrency : {1u, 4u})
                {
                    const bool rejected = shape != Shape::Shallow && shape != Shape::Deep;
                    FakeFtpEndpoint source(UploadRetention::Complete, shape != Shape::WideUnknown);
                    FakeFtpEndpoint destination(UploadRetention::Complete);
                    std::string leaf = "/selected/leaf.bin";
                    if (shape == Shape::Deep || shape == Shape::DepthLimit)
                    {
                        std::string directory = "/selected";
                        const size_t depth    = shape == Shape::Deep ? 80u : Common::FileOperations::kTraversalMaxDepth + 1u;
                        for (size_t index = 0u; index < depth; ++index)
                        {
                            directory.append("/d");
                        }
                        leaf = directory + "/leaf.bin";
                    }
                    source.SeedFile(leaf, bytes);
                    source.SeedFile("/sibling.bin", bytes);
                    if (shape == Shape::Malformed)
                    {
                        source.SetDirectoryPayload("/selected", "not a supported LIST record\r\n");
                    }
                    else if (shape == Shape::Outside)
                    {
                        source.SetDirectoryPayload("/selected", "-rw-r--r-- 1 owner group 4 Jan 01 2026 ../sibling.bin\r\n");
                    }
                    else if (shape == Shape::Overlong)
                    {
                        source.SetDirectoryPayload("/selected", std::string(CURL_MAX_WRITE_SIZE + 1u, 'x') + "\r\n");
                    }
                    else if (shape == Shape::WideUnknown)
                    {
                        // The unknown tail is after more than the queue ceiling: preflight must
                        // reach it without mutating, not reject an aggregate file-count cap.
                        std::string payload;
                        for (size_t index = 0u; index < Common::FileOperations::kTraversalMaxQueuedEntries; ++index)
                        {
                            const std::string name = std::format("a-{:05}.bin", index);
                            source.SeedFile("/selected/" + name, bytes);
                            payload.append(std::format("-rw-r--r-- 1 owner group 4 Jan 01 2026 {}\r\n", name));
                        }
                        payload.append("-rw-r--r-- 1 owner group ? Jan 01 2026 leaf.bin\r\n");
                        source.SetDirectoryPayload("/selected", std::move(payload));
                    }
                    else if (shape == Shape::StalledListCancel)
                    {
                        source.SetLateListingGate("/selected", "/never-deleted", LateListingReply::WaitForCancel);
                    }
                    destination.SeedDirectory("/destination");
                    destination.SeedFile("/sibling.bin", bytes);
                    const EndpointSnapshot sourceBefore      = source.Snapshot();
                    const EndpointSnapshot destinationBefore = destination.Snapshot();
                    HRESULT hr                               = source.Start();
                    if (SUCCEEDED(hr))
                    {
                        hr = destination.Start();
                    }
                    wil::com_ptr<IFileSystem> fileSystem;
                    if (SUCCEEDED(hr))
                    {
                        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
                    }
                    wil::com_ptr<IInformations> information;
                    if (SUCCEEDED(hr))
                    {
                        hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
                    }
                    if (SUCCEEDED(hr))
                    {
                        hr = information->SetConfiguration(
                            std::format(
                                R"({{"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":{},"deleteMaxConcurrency":1,"ftpUseEpsv":true}})",
                                concurrency)
                                .c_str());
                    }
                    FileSystemArenaOwner arena;
                    if (SUCCEEDED(hr))
                    {
                        hr = arena.Initialize(4096u);
                    }
                    if (! DebugCheck(SUCCEEDED(hr), L"native Move preflight fixture starts", *passed, *failed))
                    {
                        return E_FAIL;
                    }
                    const wchar_t* sourcePath      = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/selected", source.Port()));
                    const wchar_t* destinationPath = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/destination", destination.Port()));
                    auto* paths = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena.Get(), sizeof(const wchar_t*), alignof(const wchar_t*)));
                    if (! sourcePath || ! destinationPath || ! paths)
                    {
                        DebugCheck(false, L"Move preflight caller arena is sufficient", *passed, *failed);
                        return E_FAIL;
                    }
                    paths[0] = sourcePath;
                    CancelControl control;
                    FileSystemOptions options{};
                    options.sizeBytes           = sizeof(options);
                    options.linkPolicy          = FILESYSTEM_LINK_PRESERVE;
                    options.operationControl    = &control;
                    options.deadlineTickCount64 = GetTickCount64() + 30'000u;
                    CleanupDebtOperationCallback callback;
                    constexpr FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
                    const std::wstring beginDetail =
                        std::format(L"shape={};{};concurrency={}", static_cast<unsigned int>(shape), bulk ? L"MoveItems" : L"MoveItem", concurrency);
                    Debug::Perf::Emit(L"FileOps.Curl.MovePreflight.Begin", beginDetail.c_str(), 0u, 0u, 0u, S_OK);
                    const auto started = std::chrono::steady_clock::now();
                    std::jthread worker([&]() noexcept
                    {
                        hr = bulk ? fileSystem->MoveItems(paths, 1u, destinationPath, flags, &options, &callback, nullptr)
                                  : fileSystem->MoveItem(sourcePath, destinationPath, flags, &options, &callback, nullptr);
                    });
                    bool cooperativeCancel = false;
                    uint64_t cancelUs      = 0u;
                    if (shape == Shape::StalledListCancel)
                    {
                        const bool reached       = source.WaitForLateListingReached(3s);
                        const auto cancelStarted = std::chrono::steady_clock::now();
                        control.abortRequested.store(true, std::memory_order_release);
                        const bool returned = WaitForSingleObject(worker.native_handle(), 2'500u) == WAIT_OBJECT_0;
                        cancelUs            = Debug::Perf::ElapsedUs(cancelStarted);
                        cooperativeCancel   = reached && returned && cancelUs < 2'000'000u;
                        if (! returned)
                        {
                            source.Stop(); // Owned emergency unblock never counts as cooperative cancellation.
                        }
                    }
                    worker.join();
                    const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
                    information.reset();
                    fileSystem.reset();
                    source.Stop();
                    destination.Stop();
                    const EndpointSnapshot sourceAfter      = source.Snapshot();
                    const EndpointSnapshot destinationAfter = destination.Snapshot();
                    const HRESULT expected = shape == Shape::Malformed                                ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
                                             : shape == Shape::Outside                                ? HRESULT_FROM_WIN32(ERROR_INVALID_NAME)
                                             : shape == Shape::DepthLimit || shape == Shape::Overlong ? HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW)
                                             : shape == Shape::WideUnknown                            ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)
                                             : shape == Shape::StalledListCancel                      ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                                                                      : S_OK;
                    const size_t mutations = CountCommands(sourceAfter, "DELE") + CountCommands(sourceAfter, "RMD") + CountCommands(destinationAfter, "MKD") +
                                             CountCommands(destinationAfter, "STOR") + CountCommands(destinationAfter, "RNFR") +
                                             CountCommands(destinationAfter, "RNTO") + CountCommands(destinationAfter, "DELE") +
                                             CountCommands(destinationAfter, "RMD");
                    const std::string destinationLeaf = std::string(bulk ? "/destination/selected" : "/destination") + leaf.substr(9u);
                    const bool contents = rejected ? sourceBefore.files == sourceAfter.files && sourceBefore.directories == sourceAfter.directories &&
                                                         destinationBefore.files == destinationAfter.files &&
                                                         destinationBefore.directories == destinationAfter.directories && mutations == 0u &&
                                                         CountCommands(sourceAfter, "RETR") == 0u
                                                   : sourceAfter.files.size() == 1u && ! sourceAfter.directories.contains("/selected") &&
                                                         destinationAfter.files.size() == 2u && FileEquals(destinationAfter, destinationLeaf, bytes);
                    const bool wideReached =
                        shape != Shape::WideUnknown || CountCommands(sourceAfter, "SIZE") == Common::FileOperations::kTraversalMaxQueuedEntries + 1u;
                    const bool correct = hr == expected && SUCCEEDED(sourceAfter.serverHr) && SUCCEEDED(destinationAfter.serverHr) && contents && wideReached &&
                                         FileEquals(sourceAfter, "/sibling.bin", bytes) && FileEquals(destinationAfter, "/sibling.bin", bytes) &&
                                         (shape != Shape::StalledListCancel || cooperativeCancel);
                    const bool truthful = callback.SingleCompletionHasTruth(rejected, FileSystemRouteContract::MutationClassification::RetryableNoCommit);
                    const std::wstring route =
                        std::format(L"shape={};{};concurrency={}", static_cast<unsigned int>(shape), bulk ? L"MoveItems" : L"MoveItem", concurrency);
                    const std::wstring detail =
                        std::format(L"{};hr=0x{:08X};expected=0x{:08X};sourceServer=0x{:08X};destinationServer=0x{:08X};contents={};truthful={};wideReached={}",
                                    route,
                                    static_cast<unsigned long>(hr),
                                    static_cast<unsigned long>(expected),
                                    static_cast<unsigned long>(sourceAfter.serverHr),
                                    static_cast<unsigned long>(destinationAfter.serverHr),
                                    contents,
                                    truthful,
                                    wideReached);
                    Debug::Perf::Emit(L"FileOps.Curl.MovePreflight.TruthFixture",
                                      detail.c_str(),
                                      durationUs,
                                      mutations,
                                      CountCommands(sourceAfter, "SIZE"),
                                      correct && truthful ? S_OK : E_FAIL);
                    if (shape == Shape::StalledListCancel)
                    {
                        Debug::Perf::Emit(
                            L"FileOps.Curl.MovePreflight.ListCancel", route.c_str(), cancelUs, mutations, cooperativeCancel ? 1u : 0u, correct ? S_OK : E_FAIL);
                    }
                    DebugCheck(
                        correct, std::format(L"{} keeps exact contents and rejects unsafe discovery before any mutation", route).c_str(), *passed, *failed);
                    DebugCheck(truthful, std::format(L"{} has one truthful terminal mutation receipt", route).c_str(), *passed, *failed);
                }
            }
        }

        // The host relies on native-only Move not entering the legacy relay.
        // Invalid/non-Move modes must be rejected before provider I/O or callbacks.
        enum class MoveModeShape
        {
            NativeSameEndpoint,
            NativeCrossEndpoint,
            InvalidMode,
            NonMoveCopy,
            NonMoveDelete,
            NonMoveRename,
        };
        for (const MoveModeShape shape : {MoveModeShape::NativeSameEndpoint,
                                          MoveModeShape::NativeCrossEndpoint,
                                          MoveModeShape::InvalidMode,
                                          MoveModeShape::NonMoveCopy,
                                          MoveModeShape::NonMoveDelete,
                                          MoveModeShape::NonMoveRename})
        {
            for (const bool directory : {false, true})
            {
                for (const bool bulk : {false, true})
                {
                    for (const unsigned int concurrency : {1u, 4u})
                    {
                        const bool invalidOptions = shape != MoveModeShape::NativeSameEndpoint && shape != MoveModeShape::NativeCrossEndpoint;
                        const bool sameEndpoint   = shape != MoveModeShape::NativeCrossEndpoint && shape != MoveModeShape::InvalidMode;
                        const bool accepted       = shape == MoveModeShape::NativeSameEndpoint;
                        FakeFtpEndpoint source(UploadRetention::Complete);
                        FakeFtpEndpoint destination(UploadRetention::Complete);
                        const std::string sourceRoot = directory ? "/selected" : "/source.bin";
                        const std::string sourceLeaf = directory ? "/selected/leaf.bin" : "/source.bin";
                        const std::string targetRoot = directory ? "/destination/selected" : "/destination/source.bin";
                        const std::string targetLeaf = directory ? "/destination/selected/leaf.bin" : "/destination/source.bin";
                        source.SeedFile(sourceLeaf, bytes);
                        if (directory)
                        {
                            source.SeedFile("/selected/nested/deep.bin", bytes);
                            source.SeedDirectory("/selected/empty");
                            source.SeedFile("/selected-peer/keep.bin", bytes);
                        }
                        source.SeedFile("/sibling.bin", bytes);
                        source.SeedDirectory("/destination");
                        destination.SeedDirectory("/destination");
                        destination.SeedFile("/sibling.bin", bytes);
                        const EndpointSnapshot sourceBefore      = source.Snapshot();
                        const EndpointSnapshot destinationBefore = destination.Snapshot();
                        HRESULT hr                               = source.Start();
                        if (SUCCEEDED(hr))
                        {
                            hr = destination.Start();
                        }
                        wil::com_ptr<IFileSystem> fileSystem;
                        if (SUCCEEDED(hr))
                        {
                            hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
                        }
                        wil::com_ptr<IInformations> information;
                        if (SUCCEEDED(hr))
                        {
                            hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
                        }
                        if (SUCCEEDED(hr))
                        {
                            hr = information->SetConfiguration(
                                std::format(
                                    R"({{"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":{},"deleteMaxConcurrency":{},"ftpUseEpsv":true}})",
                                    concurrency,
                                    concurrency)
                                    .c_str());
                        }
                        FileSystemArenaOwner arena;
                        if (SUCCEEDED(hr))
                        {
                            hr = arena.Initialize(4096u);
                        }
                        if (! DebugCheck(SUCCEEDED(hr), L"Move mode fixture starts", *passed, *failed))
                        {
                            return E_FAIL;
                        }
                        const unsigned short targetPort = sameEndpoint ? source.Port() : destination.Port();
                        const wchar_t* sourcePath =
                            CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}{}", source.Port(), Utf16FromUtf8(sourceRoot)));
                        const wchar_t* targetPath =
                            CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}{}", targetPort, Utf16FromUtf8(targetRoot)));
                        const wchar_t* targetFolder = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/destination", targetPort));
                        auto* paths = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena.Get(), sizeof(const wchar_t*), alignof(const wchar_t*)));
                        const wchar_t* renameLeaf = CopyArenaString(arena.Get(), L"renamed");
                        auto* pair                = static_cast<FileSystemRenamePair*>(
                            AllocateFromFileSystemArena(arena.Get(), sizeof(FileSystemRenamePair), alignof(FileSystemRenamePair)));
                        if (! sourcePath || ! targetPath || ! targetFolder || ! paths || ! renameLeaf || ! pair)
                        {
                            DebugCheck(false, L"Move mode caller arena is sufficient", *passed, *failed);
                            return E_FAIL;
                        }
                        paths[0] = sourcePath;
                        *pair    = FileSystemRenamePair{sizeof(FileSystemRenamePair), sourcePath, renameLeaf};
                        FileSystemOptions options{};
                        options.sizeBytes = sizeof(options);
                        options.moveMode  = shape == MoveModeShape::InvalidMode ? static_cast<FileSystemMoveMode>(UINT32_MAX) : FILESYSTEM_MOVE_NATIVE_ONLY;
                        options.deadlineTickCount64 = GetTickCount64() + 30'000u;
                        CleanupDebtOperationCallback callback;
                        const FileSystemFlags flags = directory ? FILESYSTEM_FLAG_RECURSIVE : FILESYSTEM_FLAG_NONE;
                        const auto started          = std::chrono::steady_clock::now();
                        switch (shape)
                        {
                            case MoveModeShape::NonMoveCopy:
                                hr = bulk ? fileSystem->CopyItems(paths, 1u, targetFolder, flags, &options, &callback, nullptr)
                                          : fileSystem->CopyItem(sourcePath, targetPath, flags, &options, &callback, nullptr);
                                break;
                            case MoveModeShape::NonMoveDelete:
                                hr = bulk ? fileSystem->DeleteItems(paths, 1u, flags, &options, &callback, nullptr)
                                          : fileSystem->DeleteItem(sourcePath, flags, &options, &callback, nullptr);
                                break;
                            case MoveModeShape::NonMoveRename:
                                hr = bulk ? fileSystem->RenameItems(pair, 1u, flags, &options, &callback, nullptr)
                                          : fileSystem->RenameItem(sourcePath, targetPath, flags, &options, &callback, nullptr);
                                break;
                            case MoveModeShape::NativeSameEndpoint:
                            case MoveModeShape::NativeCrossEndpoint:
                            case MoveModeShape::InvalidMode:
                                hr = bulk ? fileSystem->MoveItems(paths, 1u, targetFolder, flags, &options, &callback, nullptr)
                                          : fileSystem->MoveItem(sourcePath, targetPath, flags, &options, &callback, nullptr);
                                break;
                        }
                        const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
                        information.reset();
                        fileSystem.reset();
                        source.Stop();
                        destination.Stop();
                        const EndpointSnapshot sourceAfter      = source.Snapshot();
                        const EndpointSnapshot destinationAfter = destination.Snapshot();
                        const HRESULT expected                  = invalidOptions ? E_INVALIDARG : accepted ? S_OK : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                        const size_t mutations = CountCommands(sourceAfter, "DELE") + CountCommands(sourceAfter, "RMD") + CountCommands(sourceAfter, "MKD") +
                                                 CountCommands(sourceAfter, "STOR") + CountCommands(sourceAfter, "RNFR") + CountCommands(sourceAfter, "RNTO") +
                                                 CountCommands(destinationAfter, "DELE") + CountCommands(destinationAfter, "RMD") +
                                                 CountCommands(destinationAfter, "MKD") + CountCommands(destinationAfter, "STOR") +
                                                 CountCommands(destinationAfter, "RNFR") + CountCommands(destinationAfter, "RNTO");
                        const bool noRelay     = CountCommands(sourceAfter, "RETR") == 0u && CountCommands(sourceAfter, "STOR") == 0u &&
                                                 CountCommands(destinationAfter, "RETR") == 0u && CountCommands(destinationAfter, "STOR") == 0u;
                        const bool destinationUntouched =
                            destinationAfter.files == destinationBefore.files && destinationAfter.directories == destinationBefore.directories;
                        const bool contents =
                            accepted ? ! sourceAfter.files.contains(sourceLeaf) && FileEquals(sourceAfter, targetLeaf, bytes) &&
                                           sourceAfter.files.size() == sourceBefore.files.size() &&
                                           (! directory ||
                                            (! sourceAfter.directories.contains(sourceRoot) && sourceAfter.directories.contains(targetRoot) &&
                                             ! sourceAfter.directories.contains("/selected/nested") && ! sourceAfter.directories.contains("/selected/empty") &&
                                             ! sourceAfter.files.contains("/selected/nested/deep.bin") &&
                                             sourceAfter.directories.contains("/destination/selected/nested") &&
                                             sourceAfter.directories.contains("/destination/selected/empty") &&
                                             sourceAfter.directories.size() == sourceBefore.directories.size() &&
                                             FileEquals(sourceAfter, "/destination/selected/nested/deep.bin", bytes) &&
                                             FileEquals(sourceAfter, "/selected-peer/keep.bin", bytes))) &&
                                           CountCommands(sourceAfter, "RNFR") == 1u && CountCommands(sourceAfter, "RNTO") == 1u && mutations == 2u
                                     : sourceAfter.files == sourceBefore.files && sourceAfter.directories == sourceBefore.directories && mutations == 0u;
                        const bool noInvalidIo = ! invalidOptions || (sourceAfter.commands.empty() && destinationAfter.commands.empty());
                        const bool truthful = invalidOptions ? callback.CompletedSuccessfully(0u)
                                              : accepted
                                                  ? callback.SingleCompletionHasTruth(false)
                                                  : callback.SingleCompletionHasTruth(true, FileSystemRouteContract::MutationClassification::RetryableNoCommit);
                        const bool correct  = hr == expected && contents && destinationUntouched && noRelay && noInvalidIo &&
                                              FileEquals(sourceAfter, "/sibling.bin", bytes) && FileEquals(destinationAfter, "/sibling.bin", bytes) &&
                                              SUCCEEDED(sourceAfter.serverHr) && SUCCEEDED(destinationAfter.serverHr);
                        const std::wstring detail = std::format(
                            L"shape={};directory={};bulk={};concurrency={};hr=0x{:08X};expected=0x{:08X};contents={};noRelay={};noInvalidIo={};truthful={}",
                            static_cast<unsigned int>(shape),
                            directory,
                            bulk,
                            concurrency,
                            static_cast<unsigned long>(hr),
                            static_cast<unsigned long>(expected),
                            contents,
                            noRelay,
                            noInvalidIo,
                            truthful);
                        Debug::Perf::Emit(L"FileOps.Curl.MovePreflight.ModeFixture",
                                          detail.c_str(),
                                          durationUs,
                                          mutations,
                                          sourceAfter.commands.size() + destinationAfter.commands.size(),
                                          correct && truthful ? S_OK : E_FAIL);
                        DebugCheck(correct, L"Move mode constraints preserve exact contents without an unqualified relay", *passed, *failed);
                        DebugCheck(truthful, L"Move mode rejection and success retain their actual callback truth", *passed, *failed);
                    }
                }
            }
        }

        // Two selected items exercise the real bulk scheduler, including mixed
        // native success/refusal without borrowing one item's receipt for another.
        for (const unsigned int concurrency : {1u, 4u})
        {
            FakeFtpEndpoint native(UploadRetention::Complete);
            FakeFtpEndpoint foreign(UploadRetention::Complete);
            native.SeedFile("/native.bin", bytes);
            native.SeedFile("/sibling.bin", bytes);
            native.SeedDirectory("/destination");
            foreign.SeedFile("/foreign.bin", bytes);
            const EndpointSnapshot foreignBefore = foreign.Snapshot();
            HRESULT hr                           = native.Start();
            if (SUCCEEDED(hr))
            {
                hr = foreign.Start();
            }
            wil::com_ptr<IFileSystem> fileSystem;
            if (SUCCEEDED(hr))
            {
                hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
            }
            wil::com_ptr<IInformations> information;
            if (SUCCEEDED(hr))
            {
                hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
            }
            if (SUCCEEDED(hr))
            {
                hr = information->SetConfiguration(std::format(R"({{"copyMoveMaxConcurrency":{}}})", concurrency).c_str());
            }
            FileSystemArenaOwner arena;
            if (SUCCEEDED(hr))
            {
                hr = arena.Initialize(4096u);
            }
            if (! DebugCheck(SUCCEEDED(hr), L"Mixed native Move fixture starts", *passed, *failed))
            {
                return E_FAIL;
            }
            const wchar_t* nativePath  = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/native.bin", native.Port()));
            const wchar_t* foreignPath = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/foreign.bin", foreign.Port()));
            const wchar_t* destination = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/destination", native.Port()));
            auto* paths = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena.Get(), 2u * sizeof(const wchar_t*), alignof(const wchar_t*)));
            if (! nativePath || ! foreignPath || ! destination || ! paths)
            {
                DebugCheck(false, L"Mixed native Move caller arena is sufficient", *passed, *failed);
                return E_FAIL;
            }
            paths[0] = nativePath;
            paths[1] = foreignPath;
            FileSystemOptions options{};
            options.sizeBytes           = sizeof(options);
            options.moveMode            = FILESYSTEM_MOVE_NATIVE_ONLY;
            options.deadlineTickCount64 = GetTickCount64() + 30'000u;
            CleanupDebtOperationCallback callback;
            const auto started        = std::chrono::steady_clock::now();
            hr                        = fileSystem->MoveItems(paths, 2u, destination, FILESYSTEM_FLAG_CONTINUE_ON_ERROR, &options, &callback, nullptr);
            const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
            information.reset();
            fileSystem.reset();
            native.Stop();
            foreign.Stop();
            const EndpointSnapshot nativeAfter  = native.Snapshot();
            const EndpointSnapshot foreignAfter = foreign.Snapshot();
            const size_t mutations = CountCommands(nativeAfter, "RNFR") + CountCommands(nativeAfter, "RNTO") + CountCommands(nativeAfter, "DELE") +
                                     CountCommands(nativeAfter, "STOR") + CountCommands(nativeAfter, "MKD") + CountCommands(nativeAfter, "RMD") +
                                     CountCommands(foreignAfter, "RNFR") + CountCommands(foreignAfter, "RNTO") + CountCommands(foreignAfter, "DELE") +
                                     CountCommands(foreignAfter, "STOR") + CountCommands(foreignAfter, "MKD") + CountCommands(foreignAfter, "RMD");
            const bool contents    = ! nativeAfter.files.contains("/native.bin") && nativeAfter.files.size() == 2u &&
                                     FileEquals(nativeAfter, "/destination/native.bin", bytes) && FileEquals(nativeAfter, "/sibling.bin", bytes) &&
                                     foreignAfter.files == foreignBefore.files && foreignAfter.directories == foreignBefore.directories;
            const bool noRelay     = CountCommands(nativeAfter, "RETR") == 0u && CountCommands(foreignAfter, "RETR") == 0u &&
                                     CountCommands(nativeAfter, "STOR") == 0u && CountCommands(foreignAfter, "STOR") == 0u;
            const bool truthful    = callback.TwoCompletionsKeepIndependentTruth(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
            const bool correct     = hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) && contents && noRelay && mutations == 2u &&
                                     CountCommands(nativeAfter, "RNFR") == 1u && CountCommands(nativeAfter, "RNTO") == 1u && SUCCEEDED(nativeAfter.serverHr) &&
                                     SUCCEEDED(foreignAfter.serverHr);
            const std::wstring detail = std::format(
                L"concurrency={};hr=0x{:08X};contents={};noRelay={};truthful={}", concurrency, static_cast<unsigned long>(hr), contents, noRelay, truthful);
            Debug::Perf::Emit(L"FileOps.Curl.MovePreflight.MixedModeFixture",
                              detail.c_str(),
                              durationUs,
                              mutations,
                              nativeAfter.commands.size() + foreignAfter.commands.size(),
                              correct && truthful ? S_OK : E_FAIL);
            DebugCheck(correct, L"Mixed native Move preserves the refused source without relay", *passed, *failed);
            DebugCheck(truthful, L"Mixed native Move reports independent success and no-commit refusal", *passed, *failed);
        }

        {
            // A file created after preflight must survive the source-delete pass.
            FakeFtpEndpoint source(UploadRetention::Complete);
            FakeFtpEndpoint destination(UploadRetention::Complete);
            source.SeedFile("/selected/kept.bin", bytes);
            source.SeedFileAfterNextRetrieve("/selected/late.bin", bytes);
            destination.SeedDirectory("/destination");
            HRESULT hr = source.Start();
            if (SUCCEEDED(hr))
            {
                hr = destination.Start();
            }
            wil::com_ptr<IFileSystem> fileSystem;
            if (SUCCEEDED(hr))
            {
                hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
            }
            FileSystemArenaOwner arena;
            if (SUCCEEDED(hr))
            {
                hr = arena.Initialize(4096u);
            }
            const wchar_t* sourcePath      = SUCCEEDED(hr) ? CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/selected", source.Port())) : nullptr;
            const wchar_t* destinationPath = SUCCEEDED(hr) ? CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/destination", destination.Port())) : nullptr;
            if (! DebugCheck(SUCCEEDED(hr) && sourcePath && destinationPath, L"late-writer Move fixture starts", *passed, *failed))
            {
                return E_FAIL;
            }
            CleanupDebtOperationCallback callback;
            FileSystemOptions options{};
            options.sizeBytes        = sizeof(options);
            options.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
            options.operationControl = nullptr;
            const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
            hr = fileSystem->MoveItem(sourcePath, destinationPath, flags, &options, &callback, nullptr);
            fileSystem.reset();
            source.Stop();
            destination.Stop();
            const EndpointSnapshot sourceAfter      = source.Snapshot();
            const EndpointSnapshot destinationAfter = destination.Snapshot();
            const bool latePreserved                = FileEquals(sourceAfter, "/selected/late.bin", bytes);
            const bool lateNotDeleted               = std::ranges::none_of(sourceAfter.deletedPaths, [](const std::string& path) noexcept { return path == "/selected/late.bin"; });
            const bool copiedKept                   = FileEquals(destinationAfter, "/destination/kept.bin", bytes);
            const bool lateNotCopied                = ! destinationAfter.files.contains("/destination/late.bin");
            DebugCheck(latePreserved && lateNotDeleted && copiedKept && lateNotCopied,
                       L"Move source-delete keeps a file created after preflight",
                       *passed,
                       *failed);
            DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                       L"Move reports partial when a late source file blocks source cleanup",
                       *passed,
                       *failed);
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: report fixture failure without unwinding into the host.
        Debug::Error(L"Curl Move preflight fixture failed after std::exception.");
        DebugCheck(false, L"Move preflight fixture must not throw", *passed, *failed);
    }
    return *failed == 0u ? S_OK : E_FAIL;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlEntryLookupTruthForSelfTest(unsigned int* passed, unsigned int* failed) noexcept
{
    if (! passed || ! failed)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"entry lookup fixture initializes Winsock", *passed, *failed))
    {
        return E_FAIL;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    try
    {
        for (const LoopbackEndpointMode mode : {LoopbackEndpointMode::Listening, LoopbackEndpointMode::BoundOnly})
        {
            wil::unique_socket endpoint;
            unsigned short port      = 0u;
            const HRESULT endpointHr = CreateLoopbackEndpoint(endpoint, port, mode);
            BOOL exclusive           = FALSE;
            int optionBytes          = static_cast<int>(sizeof(exclusive));
            const bool ownsPort      = SUCCEEDED(endpointHr) && port != 0u &&
                                       getsockopt(endpoint.get(), SOL_SOCKET, SO_EXCLUSIVEADDRUSE, reinterpret_cast<char*>(&exclusive), &optionBytes) == 0 &&
                                       optionBytes == static_cast<int>(sizeof(exclusive)) && exclusive != FALSE;
            wil::unique_socket challenger(socket(AF_INET, SOCK_STREAM, IPPROTO_TCP));
            const BOOL reuse = TRUE;
            bool rejected    = false;
            int bindError    = 0;
            if (ownsPort && challenger &&
                setsockopt(challenger.get(), SOL_SOCKET, SO_REUSEADDR, reinterpret_cast<const char*>(&reuse), static_cast<int>(sizeof(reuse))) == 0)
            {
                sockaddr_in address{};
                address.sin_family      = AF_INET;
                address.sin_addr.s_addr = htonl(INADDR_LOOPBACK);
                address.sin_port        = htons(port);
                const int bindResult    = bind(challenger.get(), reinterpret_cast<const sockaddr*>(&address), static_cast<int>(sizeof(address)));
                bindError               = bindResult == 0 ? 0 : WSAGetLastError();
                rejected                = bindResult == SOCKET_ERROR && (bindError == WSAEACCES || bindError == WSAEADDRINUSE);
            }
            const bool correct    = ownsPort && rejected;
            const wchar_t* detail = mode == LoopbackEndpointMode::Listening ? L"listening" : L"bound-only";
            Debug::Perf::Emit(L"FileOps.Curl.FakeFtp.ExclusiveEndpoint", detail, 0u, port, static_cast<uint64_t>(bindError), correct ? S_OK : E_FAIL);
            DebugCheck(correct, L"owned loopback endpoint is exclusive and rejects a competing reuse bind", *passed, *failed);
        }
        {
            const int priorError = WSAGetLastError();
            auto restoreError    = wil::scope_exit([priorError]() noexcept { WSASetLastError(priorError); });
            WSASetLastError(WSAENETDOWN);
            unsigned short port       = 0u;
            const HRESULT probeHr     = ProbeCurlEphemeralBindForSelfTest(port);
            const int retainedError   = WSAGetLastError();
            const bool correct        = probeHr == S_OK && port != 0u && retainedError == WSAENETDOWN;
            const std::wstring detail = std::format(L"control;port={};retainedError={}", port, retainedError);
            Debug::Perf::Emit(L"FileOps.Curl.ConnectionAllocationProbe", detail.c_str(), 0u, 0u, port, probeHr);
            DebugCheck(correct, L"fresh ephemeral bind succeeds without changing the caller's Winsock error", *passed, *failed);
        }
        for (const bool bindFirst : {false, true})
        {
            const int priorError = WSAGetLastError();
            auto restoreError    = wil::scope_exit([priorError]() noexcept { WSASetLastError(priorError); });
            WSASetLastError(WSAENETDOWN);
            unsigned short port       = 0u;
            const wchar_t* stage      = nullptr;
            const auto started        = std::chrono::steady_clock::now();
            const HRESULT probeHr     = ProbeCurlLoopbackConnectForSelfTest(bindFirst, port, stage);
            const int retainedError   = WSAGetLastError();
            const bool correct        = probeHr == S_OK && port != 0u && std::wstring_view(stage) == L"connected" && retainedError == WSAENETDOWN;
            const std::wstring detail = std::format(L"phase=control;explicitBind={};stage={};retainedError={}", bindFirst, stage, retainedError);
            Debug::Perf::Emit(L"FileOps.Curl.LoopbackConnectProbe", detail.c_str(), Debug::Perf::ElapsedUs(started), bindFirst ? 1u : 0u, port, probeHr);
            DebugCheck(correct, L"fresh owned loopback connect succeeds and preserves Winsock error", *passed, *failed);
        }
        {
            struct TraceCase final
            {
                std::wstring_view name;
                curl_infotype type;
                std::string_view text;
                bool accepted;
                int remote;
                int local;
            };
            constexpr std::array<TraceCase, 18u> cases{{
                {L"loopback", CURLINFO_TEXT, "connect to 127.0.0.1 port 50000 from 127.0.0.1 port 51000 failed: error", true, 50000, 51000},
                {L"wildcard", CURLINFO_TEXT, "connect to 127.0.0.1 port 21 from 0.0.0.0 port 0 failed: error", true, 21, 0},
                {L"unavailable", CURLINFO_TEXT, "connect to 127.0.0.1 port 21 from  port -1 failed: error", true, 21, -1},
                {L"limits", CURLINFO_TEXT, "connect to 127.0.0.1 port 65535 from 127.0.0.1 port 65535 failed: error", true, 65535, 65535},
                {L"discard-suffix", CURLINFO_TEXT, "connect to 127.0.0.1 port 21 from  port 0 failed: synthetic-private-sentinel", true, 21, 0},
                {L"empty", CURLINFO_TEXT, "", false, 0, 0},
                {L"header-in", CURLINFO_HEADER_IN, "connect to 127.0.0.1 port 21 from  port 0 failed: error", false, 0, 0},
                {L"header-out", CURLINFO_HEADER_OUT, "connect to 127.0.0.1 port 21 from  port 0 failed: error", false, 0, 0},
                {L"payload", CURLINFO_DATA_IN, "connect to 127.0.0.1 port 21 from  port 0 failed: error", false, 0, 0},
                {L"unrelated", CURLINFO_TEXT, "USER synthetic-private-sentinel", false, 0, 0},
                {L"other-peer", CURLINFO_TEXT, "connect to 192.0.2.1 port 21 from  port 0 failed: error", false, 0, 0},
                {L"remote-zero", CURLINFO_TEXT, "connect to 127.0.0.1 port 0 from  port 0 failed: error", false, 0, 0},
                {L"remote-overflow", CURLINFO_TEXT, "connect to 127.0.0.1 port 65536 from  port 0 failed: error", false, 0, 0},
                {L"local-overflow", CURLINFO_TEXT, "connect to 127.0.0.1 port 21 from  port 65536 failed: error", false, 0, 0},
                {L"local-negative", CURLINFO_TEXT, "connect to 127.0.0.1 port 21 from  port -2 failed: error", false, 0, 0},
                {L"signed-remote", CURLINFO_TEXT, "connect to 127.0.0.1 port +21 from  port 0 failed: error", false, 0, 0},
                {L"truncated", CURLINFO_TEXT, "connect to 127.0.0.1 port 21 from  port 0", false, 0, 0},
                {L"numeric-tail", CURLINFO_TEXT, "connect to 127.0.0.1 port 21 from  port 0x failed: error", false, 0, 0},
            }};
            for (const TraceCase& fixture : cases)
            {
                const auto ports   = CurlDirectoryCursor::ParseFailedConnectTraceForSelfTest(fixture.type, fixture.text);
                const bool correct = ports.has_value() == fixture.accepted &&
                                     (! ports.has_value() || (ports.value().remote == fixture.remote && ports.value().local == fixture.local));
                Debug::Perf::Emit(L"FileOps.Curl.SocketConnectFailure.Control",
                                  fixture.name.data(),
                                  0u,
                                  fixture.accepted ? 1u : 0u,
                                  ports.has_value() ? 1u : 0u,
                                  correct ? S_OK : E_FAIL);
                DebugCheck(correct, L"connection trace retains only reviewed numeric loopback fields", *passed, *failed);
            }
        }
        {
            FakeFtpEndpoint known(UploadRetention::Complete);
            FakeFtpEndpoint unknown(UploadRetention::Complete, true, false);
            for (FakeFtpEndpoint* endpoint : std::array<FakeFtpEndpoint*, 2u>{{&known, &unknown}})
            {
                endpoint->SeedDirectory("/empty");
                endpoint->SeedDirectory("/chosen/sub/deeper");
                endpoint->SeedFile("/root.bin", {1u, 2u, 3u, 4u});
                endpoint->SeedFile("/chosen/a.bin", {});
                endpoint->SeedFile("/chosen/z literal%20 -> \xC3\xA9.bin", {1u, 2u, 3u});
                endpoint->SeedFile("/chosen/sub/hidden.bin", {1u, 2u, 3u, 4u});
                endpoint->SeedFile("/chosen-prefix/outside.bin", {1u});
            }
            const auto checkPayload =
                [&](const wchar_t* label, const FakeFtpEndpoint& endpoint, std::string_view directory, bool namesOnly, std::string_view expected)
            {
                const auto started        = std::chrono::steady_clock::now();
                const std::string payload = endpoint.BuildDirectoryPayload(directory, namesOnly);
                const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
                const bool correct        = payload == expected;
                Debug::Perf::Emit(L"FileOps.Curl.FixtureListing.Bytes", label, durationUs, payload.size(), expected.size(), correct ? S_OK : E_FAIL);
                DebugCheck(correct, std::format(L"FTP fixture listing preserves literal bytes: {}", label).c_str(), *passed, *failed);
            };
            // Hand-authored expected bytes independently protect ordering,
            // direct-child membership, unknown sizes and literal UTF-8 names.
            constexpr std::string_view knownListing = "drwxr-xr-x 1 owner group 0 Jan 01 2026 sub\r\n"
                                                      "-rw-r--r-- 1 owner group 0 Jan 01 2026 a.bin\r\n"
                                                      "-rw-r--r-- 1 owner group 3 Jan 01 2026 z literal%20 -> \xC3\xA9.bin\r\n";
            checkPayload(L"empty", known, "/empty", false, "");
            checkPayload(L"root-names", known, "/", true, "chosen\r\nchosen-prefix\r\nempty\r\nroot.bin\r\n");
            checkPayload(L"selected-list", known, "/chosen", false, knownListing);
            checkPayload(L"selected-names", known, "/chosen", true, "sub\r\na.bin\r\nz literal%20 -> \xC3\xA9.bin\r\n");
            checkPayload(L"nested-list",
                         known,
                         "/chosen/sub",
                         false,
                         "drwxr-xr-x 1 owner group 0 Jan 01 2026 deeper\r\n"
                         "-rw-r--r-- 1 owner group 4 Jan 01 2026 hidden.bin\r\n");
            checkPayload(L"missing", known, "/missing", false, "");
            checkPayload(L"unknown-size",
                         unknown,
                         "/chosen",
                         false,
                         "drwxr-xr-x 1 owner group 0 Jan 01 2026 sub\r\n"
                         "-rw-r--r-- 1 owner group ? Jan 01 2026 a.bin\r\n"
                         "-rw-r--r-- 1 owner group ? Jan 01 2026 z literal%20 -> \xC3\xA9.bin\r\n");
            known.SetListedSize("/chosen/a.bin", 123456789u);
            checkPayload(L"size-override",
                         known,
                         "/chosen",
                         false,
                         "drwxr-xr-x 1 owner group 0 Jan 01 2026 sub\r\n"
                         "-rw-r--r-- 1 owner group 123456789 Jan 01 2026 a.bin\r\n"
                         "-rw-r--r-- 1 owner group 3 Jan 01 2026 z literal%20 -> \xC3\xA9.bin\r\n");
            known.SeedFile("/chosen/b.bin", {1u});
            checkPayload(L"new-generation",
                         known,
                         "/chosen",
                         false,
                         "drwxr-xr-x 1 owner group 0 Jan 01 2026 sub\r\n"
                         "-rw-r--r-- 1 owner group 123456789 Jan 01 2026 a.bin\r\n"
                         "-rw-r--r-- 1 owner group 1 Jan 01 2026 b.bin\r\n"
                         "-rw-r--r-- 1 owner group 3 Jan 01 2026 z literal%20 -> \xC3\xA9.bin\r\n");
            constexpr std::string_view malformed{"bad\0row\n", 8u};
            known.SetDirectoryPayload("/chosen", std::string(malformed));
            checkPayload(L"custom-list", known, "/chosen", false, malformed);
            checkPayload(L"custom-names", known, "/chosen", true, malformed);

            FakeFtpEndpoint wide(UploadRetention::Complete);
            std::string expectedWide;
            constexpr size_t fileCount = Common::FileOperations::kTraversalMaxQueuedEntries + 1u;
            constexpr size_t repeats   = 64u;
            for (size_t index = 0u; index < fileCount; ++index)
            {
                const std::string leaf = std::format("file-{:04}.bin", index);
                wide.SeedFile(std::format("/wide/{}", leaf), {1u, 2u, 3u});
                expectedWide.append("-rw-r--r-- 1 owner group 3 Jan 01 2026 ");
                expectedWide.append(leaf);
                expectedWide.append("\r\n");
            }
            bool exactBytes    = true;
            size_t bytes       = 0u;
            const auto started = std::chrono::steady_clock::now();
            for (size_t repeat = 0u; repeat < repeats; ++repeat)
            {
                const std::string payload = wide.BuildDirectoryPayload("/wide", false);
                exactBytes                = (payload == expectedWide) && exactBytes;
                bytes += payload.size();
            }
            const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
            Debug::Perf::Emit(L"FileOps.Curl.FixtureListing.Wide", L"files=4097;renders=64", durationUs, repeats, bytes, exactBytes ? S_OK : E_FAIL);
            DebugCheck(
                exactBytes && bytes == expectedWide.size() * repeats, L"FTP fixture repeatedly serializes all 4097 files with exact bytes", *passed, *failed);
        }
        {
            // Parent/child lifetime witness: retain cursor objects and unread
            // parent records while their completed transports return to the pool.
            FakeFtpEndpoint endpoint(UploadRetention::Complete);
            endpoint.SeedFile("/parent/child/leaf.bin", {1u, 2u, 3u, 4u});
            endpoint.SeedFile("/parent/tail.bin", {5u, 6u, 7u});
            endpoint.SeedDirectory("/empty");
            endpoint.SeedDirectory("/wide");
            std::string widePayload;
            constexpr size_t wideRows = 4'097u;
            for (size_t index = 0u; index < wideRows; ++index)
            {
                widePayload += std::format("-rw-r--r-- 1 owner group 1 Jan 01 2026 row-{:04}.bin\r\n", index);
            }
            endpoint.SetDirectoryPayload("/wide", widePayload);
            HRESULT hr = endpoint.Start();
            ResolvedLocation resolved{};
            if (SUCCEEDED(hr))
            {
                hr = ResolveLocation(
                    Protocol::Ftp, FileSystemCurl::Settings{}, std::format(L"//anonymous@127.0.0.1:{}/parent", endpoint.Port()), nullptr, true, resolved);
            }
            if (! DebugCheck(SUCCEEDED(hr), L"nested cursor lifetime fixture starts", *passed, *failed))
            {
                return E_FAIL;
            }
            const auto observe = [&](const wchar_t* name, bool correct, uint64_t value0 = 0u, uint64_t value1 = 0u)
            {
                Debug::Perf::Emit(L"FileOps.Curl.CursorLifetime", name, 0u, value0, value1, correct ? S_OK : E_FAIL);
                DebugCheck(correct, name, *passed, *failed);
            };
            const auto checkpoint = []() noexcept { return S_OK; };
            CurlDirectoryCursor parent;
            FilesInformationCurl::Entry entry{};
            hr = parent.Open(resolved.connection, L"/parent", checkpoint);
            if (SUCCEEDED(hr))
            {
                hr = parent.Next(entry);
            }
            observe(L"parent releases completed transfer before yielding child",
                    hr == S_OK && entry.name == L"child" && (entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && ! parent.HasActiveTransferForSelfTest());

            CurlDirectoryCursor child;
            hr = child.Open(resolved.connection, L"/parent/child", checkpoint);
            if (SUCCEEDED(hr))
            {
                hr = child.Next(entry);
            }
            const bool childCorrect = hr == S_OK && entry.name == L"leaf.bin" && entry.sizeKnown && entry.sizeBytes == 4u &&
                                      ! child.HasActiveTransferForSelfTest() && child.Next(entry) == S_FALSE;
            observe(L"nested child reuses completed transport", childCorrect);
            hr                       = parent.Next(entry);
            const bool parentCorrect = hr == S_OK && entry.name == L"tail.bin" && entry.sizeKnown && entry.sizeBytes == 3u && parent.Next(entry) == S_FALSE &&
                                       ! parent.HasActiveTransferForSelfTest();
            observe(L"parent buffered records survive child reuse", parentCorrect);
            const EndpointSnapshot nested = endpoint.Snapshot();
            const size_t sessions         = CountCommands(nested, "USER");
            const size_t listings         = CountCommands(nested, "LIST");
            observe(L"nested small listings use one authenticated session", sessions == 1u && listings == 2u, sessions, listings);

            CurlDirectoryCursor empty;
            hr = empty.Open(resolved.connection, L"/empty", checkpoint);
            observe(L"empty listing releases transfer", SUCCEEDED(hr) && empty.Next(entry) == S_FALSE && ! empty.HasActiveTransferForSelfTest());

            CurlDirectoryCursor wide;
            hr             = wide.Open(resolved.connection, L"/wide", checkpoint);
            size_t rows    = 0u;
            bool exactRows = true;
            bool streamed  = false;
            if (SUCCEEDED(hr))
            {
                while ((hr = wide.Next(entry)) == S_OK)
                {
                    streamed  = streamed || wide.HasActiveTransferForSelfTest();
                    exactRows = exactRows && entry.name == std::format(L"row-{:04}.bin", rows) && entry.sizeKnown && entry.sizeBytes == 1u;
                    ++rows;
                }
            }
            observe(L"multi-buffer listing preserves all rows then releases transfer",
                    hr == S_FALSE && rows == wideRows && exactRows && streamed && ! wide.HasActiveTransferForSelfTest(),
                    rows,
                    wideRows);

            endpoint.SeedDirectory("/wide-child/nested");
            std::string wideChildPayload = "drwxr-xr-x 1 owner group 0 Jan 01 2026 nested\r\n";
            wideChildPayload += widePayload;
            endpoint.SetDirectoryPayload("/wide-child", std::move(wideChildPayload));
            CurlDirectoryCursor wideChild;
            hr                    = wideChild.Open(resolved.connection, L"/wide-child", checkpoint);
            bool firstIsDirectory = false;
            bool pausedOnFirst    = false;
            if (SUCCEEDED(hr))
            {
                hr                = wideChild.Next(entry);
                firstIsDirectory  = hr == S_OK && entry.name == L"nested" && (entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                pausedOnFirst     = firstIsDirectory && wideChild.HasActiveTransferForSelfTest();
                while ((hr = wideChild.Next(entry)) == S_OK)
                {
                }
            }
            observe(L"paused listing still holds transfer when the first row is a directory", pausedOnFirst);
            observe(L"finishing that listing releases the parent before a child Open is safe",
                    hr == S_FALSE && firstIsDirectory && ! wideChild.HasActiveTransferForSelfTest());

            bool abandonedActive = false;
            {
                CurlDirectoryCursor abandoned;
                hr              = abandoned.Open(resolved.connection, L"/wide", checkpoint);
                abandonedActive = SUCCEEDED(hr) && abandoned.Next(entry) == S_OK && abandoned.HasActiveTransferForSelfTest();
            }
            hr = GetEntryInfo(resolved.connection, L"/parent/tail.bin", entry);
            observe(L"abandoned paused cursor permits safe reborrow", abandonedActive && hr == S_OK && entry.name == L"tail.bin");

            bool cancelled = false;
            CurlDirectoryCursor cancel;
            hr        = cancel.Open(resolved.connection, L"/parent", [&]() noexcept { return cancelled ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK; });
            cancelled = true;
            const bool canceledCorrect = SUCCEEDED(hr) && cancel.Next(entry) == HRESULT_FROM_WIN32(ERROR_CANCELLED) && ! cancel.HasActiveTransferForSelfTest();
            hr                         = GetEntryInfo(resolved.connection, L"/parent/tail.bin", entry);
            observe(L"cancelled cursor returns borrow without poisoning next lookup", canceledCorrect && hr == S_OK && entry.name == L"tail.bin");

            endpoint.SetFailedListFinalReply(true);
            CurlDirectoryCursor failedCursor;
            entry = {};
            hr    = failedCursor.Open(resolved.connection, L"/parent", checkpoint);
            if (SUCCEEDED(hr))
            {
                hr = failedCursor.Next(entry);
            }
            const bool failedCorrect = FAILED(hr) && entry.name.empty() && ! failedCursor.HasActiveTransferForSelfTest();
            hr                       = GetEntryInfo(resolved.connection, L"/parent/tail.bin", entry);
            observe(L"bad final reply publishes no small-list row or lookup result", failedCorrect && FAILED(hr) && entry.name.empty());
            endpoint.SetFailedListFinalReply(false);
            hr = GetEntryInfo(resolved.connection, L"/parent/tail.bin", entry);
            endpoint.Stop();
            observe(L"failed completed cursor permits healthy reborrow",
                    hr == S_OK && entry.name == L"tail.bin" && entry.sizeKnown && entry.sizeBytes == 3u && SUCCEEDED(endpoint.Snapshot().serverHr));
        }
        enum class Shape
        {
            Literal,
            Missing,
            MalformedBefore,
            MalformedAfter,
            Duplicate,
            Outside,
            Overlong,
            Cancel
        };
        {
            wil::unique_socket refusedEndpoint;
            unsigned short port = 0u;
            HRESULT hr          = CreateLoopbackEndpoint(refusedEndpoint, port, LoopbackEndpointMode::BoundOnly);
            ResolvedLocation resolved{};
            if (SUCCEEDED(hr))
            {
                hr = ResolveLocation(
                    Protocol::Ftp, FileSystemCurl::Settings{}, std::format(L"//anonymous@127.0.0.1:{}/never-created.bin", port), nullptr, true, resolved);
            }
            if (! DebugCheck(SUCCEEDED(hr), L"refused endpoint keeps exclusive test-port ownership", *passed, *failed))
            {
                return E_FAIL;
            }
            // Windows may take about two seconds to finish a refused loopback
            // connect. Let this refusal oracle observe that result, not the
            // shorter timeout; production and traversal-test limits are unchanged.
            resolved.connection.connectTimeoutMs   = 5'000u;
            resolved.connection.operationTimeoutMs = 7'000u;
            constexpr HRESULT refused              = HRESULT_FROM_WIN32(ERROR_CONNECTION_REFUSED);
            const auto started                     = std::chrono::steady_clock::now();
            FilesInformationCurl::Entry entry{};
            const HRESULT lookupHr   = GetEntryInfo(resolved.connection, resolved.remotePath, entry);
            const bool lookupCorrect = lookupHr == refused && entry.name.empty();
            DebugCheck(lookupCorrect, L"refused cursor returns failure without publishing an entry", *passed, *failed);
            {
                CurlDirectoryCursor cursor;
                const bool initiallyEmpty = ! cursor.FailedConnectPortsForSelfTest().has_value();
                HRESULT cursorHr          = cursor.Open(resolved.connection, L"/", []() noexcept { return S_OK; });
                if (SUCCEEDED(cursorHr))
                {
                    cursorHr = cursor.Next(entry);
                }
                const auto ports   = cursor.FailedConnectPortsForSelfTest();
                const bool correct = initiallyEmpty && cursorHr == refused && ports.has_value() && ports.value().remote == port;
                Debug::Perf::Emit(L"FileOps.Curl.SocketConnectFailure.Control",
                                  L"actual-refused-cursor",
                                  0u,
                                  port,
                                  ports.has_value() ? static_cast<uint64_t>(ports.value().remote) : 0u,
                                  correct ? S_OK : E_FAIL);
                DebugCheck(correct, L"actual refused cursor captures the owned endpoint port through libcurl diagnostics", *passed, *failed);
            }
            uint64_t size           = 1u;
            bool known              = true;
            const HRESULT probeHr   = CurlProbeRemoteFileSize(resolved.connection, resolved.remotePath, size, known);
            const bool probeCorrect = probeHr == refused && size == 0u && ! known;
            DebugCheck(probeCorrect, L"refused size probe remains unknown after bounded read attempts", *passed, *failed);
            const HRESULT quoteHr   = CurlPerformQuote(resolved.connection, {"DELE /never-created.bin"});
            const bool quoteCorrect = quoteHr == refused;
            DebugCheck(quoteCorrect, L"refused namespace connection returns failure without replay", *passed, *failed);
            const std::wstring detail = std::format(L"bound-not-listening;cursor=0x{:08X};size-probe=0x{:08X};quote=0x{:08X}",
                                                    static_cast<unsigned long>(lookupHr),
                                                    static_cast<unsigned long>(probeHr),
                                                    static_cast<unsigned long>(quoteHr));
            Debug::Perf::Emit(L"FileOps.Curl.EntryLookup.RefusedEndpoint",
                              detail.c_str(),
                              Debug::Perf::ElapsedUs(started),
                              3u,
                              port,
                              lookupCorrect && probeCorrect && quoteCorrect ? S_OK : E_FAIL);
        }
        {
            struct FramingCase
            {
                std::wstring_view name;
                std::string payload;
                HRESULT expected;
                std::vector<std::wstring> names;
            };
            constexpr size_t framingLimit   = static_cast<size_t>(CURL_MAX_WRITE_SIZE);
            const std::wstring framingName  = L"literal%20 -> \u00E9.bin";
            const std::string framingRecord = "-rw-r--r-- 1 owner group 4 Jan 01 2026 literal%20 -> \xC3\xA9.bin\r\n";
            const std::string maxHeader     = std::format("total {}0", std::string(framingLimit - 7u, ' '));
            const HRESULT overflow          = HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
            const HRESULT invalid           = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            const std::array<FramingCase, 12u> framingCases{{
                {L"empty", {}, S_OK, {}},
                {L"mixed",
                 "\r\ntotal 0\r\ndrwxr-xr-x 1 owner group 0 Jan 01 2026 .\r\n" + framingRecord + "01-02-26  03:04PM  5 second.bin",
                 S_OK,
                 {framingName, L"second.bin"}},
                {L"limit-lf", maxHeader + "\n" + framingRecord, S_OK, {framingName}},
                {L"limit-crlf", std::format("total {}0\r\n", std::string(framingLimit - 8u, ' ')) + framingRecord, S_OK, {framingName}},
                {L"limit-eof", maxHeader, S_OK, {}},
                {L"overflow-lf", maxHeader + " \n", overflow, {}},
                {L"overflow-eof", maxHeader + " ", overflow, {}},
                {L"invalid-utf8", "-rw-r--r-- 1 owner group 4 Jan 01 2026 \xC3\x28.bin\r\n", invalid, {}},
                {L"embedded-nul", framingRecord + std::string(1u, '\0') + "\n", invalid, {framingName}},
                {L"malformed-after", framingRecord + "invalid LIST record\n", invalid, {framingName}},
                {L"dos-size-suffix", "01-02-26  03:04PM  4junk trailing.bin\r\n", invalid, {}},
                {L"dos-size-fraction", "01-02-26  03:04PM  4.5 decimal.bin\r\n", invalid, {}},
            }};
            // Fixture-specific oracle: names/sizes are authored facts, not output
            // from a second parser that could share the same framing mistake.
            const auto recordsMatch = [](const FramingCase& fixture, HRESULT hr, const std::vector<FilesInformationCurl::Entry>& entries) noexcept
            {
                if (hr != fixture.expected || entries.size() != fixture.names.size())
                {
                    return false;
                }
                for (size_t index = 0u; index < entries.size(); ++index)
                {
                    const auto& entry           = entries[index];
                    const uint64_t expectedSize = fixture.names[index] == L"second.bin" ? 5u : 4u;
                    if (entry.name != fixture.names[index] || ! entry.sizeKnown || entry.sizeBytes != expectedSize || entry.lastWriteTime == 0)
                    {
                        return false;
                    }
                }
                return true;
            };
            for (const FramingCase& fixture : framingCases)
            {
                for (const size_t chunkWidth : std::array<size_t, 3u>{{1u, 7u, framingLimit}})
                {
                    std::vector<std::string_view> chunks;
                    for (size_t offset = 0u; offset < fixture.payload.size(); offset += (std::min)(chunkWidth, fixture.payload.size() - offset))
                    {
                        chunks.push_back(std::string_view(fixture.payload).substr(offset, chunkWidth));
                    }
                    std::vector<FilesInformationCurl::Entry> entries;
                    const auto started        = std::chrono::steady_clock::now();
                    const HRESULT hr          = CurlDirectoryCursor::ParseChunksForSelfTest(chunks, entries);
                    const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
                    const bool correct        = recordsMatch(fixture, hr, entries);
                    const std::wstring detail = std::format(L"{};chunkWidth={};hr=0x{:08X};expected=0x{:08X}",
                                                            fixture.name,
                                                            chunkWidth,
                                                            static_cast<unsigned long>(hr),
                                                            static_cast<unsigned long>(fixture.expected));
                    Debug::Perf::Emit(
                        L"FileOps.Curl.EntryLookup.Framing", detail.c_str(), durationUs, fixture.payload.size(), chunks.size(), correct ? S_OK : E_FAIL);
                    DebugCheck(correct, std::format(L"cursor framing preserves exact rows and bounds: {}", detail).c_str(), *passed, *failed);
                }
            }
            {
                const FramingCase& fixture = framingCases[1u];
                std::vector<FilesInformationCurl::Entry> entries;
                std::optional<size_t> firstFailure;
                const auto started = std::chrono::steady_clock::now();
                for (size_t split = 0u; split <= fixture.payload.size(); ++split)
                {
                    const std::array<std::string_view, 2u> chunks{std::string_view(fixture.payload).substr(0u, split),
                                                                  std::string_view(fixture.payload).substr(split)};
                    const HRESULT hr = CurlDirectoryCursor::ParseChunksForSelfTest(chunks, entries);
                    if (! recordsMatch(fixture, hr, entries) && ! firstFailure.has_value())
                    {
                        firstFailure = split;
                    }
                }
                const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
                const std::wstring detail = std::format(L"all-two-chunk-splits;attempts={};firstFailure={}",
                                                        fixture.payload.size() + 1u,
                                                        firstFailure.value_or((std::numeric_limits<size_t>::max)()));
                Debug::Perf::Emit(L"FileOps.Curl.EntryLookup.Framing",
                                  detail.c_str(),
                                  durationUs,
                                  fixture.payload.size(),
                                  2u * (fixture.payload.size() + 1u),
                                  firstFailure.has_value() ? E_FAIL : S_OK);
                DebugCheck(! firstFailure.has_value(), std::format(L"cursor handles every CR/LF/UTF-8 split: {}", detail).c_str(), *passed, *failed);
            }
            {
                const std::string oversized(framingLimit + 1u, 'x');
                const std::array<std::string_view, 1u> chunks{oversized};
                std::vector<FilesInformationCurl::Entry> entries;
                const auto started = std::chrono::steady_clock::now();
                const HRESULT hr   = CurlDirectoryCursor::ParseChunksForSelfTest(chunks, entries);
                const bool correct = hr == overflow && entries.empty();
                Debug::Perf::Emit(L"FileOps.Curl.EntryLookup.Framing",
                                  L"oversized-callback",
                                  Debug::Perf::ElapsedUs(started),
                                  oversized.size(),
                                  chunks.size(),
                                  correct ? S_OK : E_FAIL);
                DebugCheck(correct, L"cursor rejects an oversized transport callback before copying bytes", *passed, *failed);
            }
            {
                const std::array<std::pair<std::string_view, std::string_view>, 2u> sockets{{
                    {"srwxrwxrwx 1 1000 1000 0 Jan 01 12:00 mysql.sock\r\n", "numeric-owner"},
                    {"srwxrwxrwx 1 mysql mysql 0 Jan 01 12:00 mysql.sock\r\n", "named-owner"},
                }};
                bool socketsCorrect = true;
                for (const auto& [payload, _] : sockets)
                {
                    const std::array<std::string_view, 1u> chunks{payload};
                    std::vector<FilesInformationCurl::Entry> entries;
                    const HRESULT hr = CurlDirectoryCursor::ParseChunksForSelfTest(chunks, entries);
                    socketsCorrect   = socketsCorrect && hr == S_OK && entries.size() == 1u && entries.front().name == L"mysql.sock" &&
                                     (entries.front().attributes & FILE_ATTRIBUTE_DEVICE) != 0 && ! entries.front().sizeKnown;
                }
                const std::array<std::string_view, 1u> phantom{"not-date  03:04PM                4 phantom.bin\r\n"};
                std::vector<FilesInformationCurl::Entry> phantomEntries;
                const HRESULT phantomHr = CurlDirectoryCursor::ParseChunksForSelfTest(phantom, phantomEntries);
                DebugCheck(socketsCorrect, L"Unix socket/device rows stay Unix names, never DOS phantoms", *passed, *failed);
                DebugCheck(phantomHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && phantomEntries.empty(),
                           L"DOS listing requires a real date/time, not a numeric third token",
                           *passed,
                           *failed);
            }
        }
        // Same cursor callback/framing/parser as native lookup, without FTP,
        // server generation or whole-directory result retention in the timed loop.
        for (const bool dos : {false, true})
        {
            std::string listing;
            constexpr size_t rowCount = 4'097u;
            constexpr size_t repeats  = 64u;
            for (size_t index = 0u; index < rowCount; ++index)
            {
                listing.append(dos ? std::format("01-02-26  03:04PM  4 f-{:05}.bin\r\n", index)
                                   : std::format("-rw-r--r-- 1 owner group 4 Jan 02 2026 f-{:05}.bin\r\n", index));
            }
            std::vector<std::string_view> chunks;
            for (size_t offset = 0u; offset < listing.size(); offset += (std::min)(static_cast<size_t>(CURL_MAX_WRITE_SIZE), listing.size() - offset))
            {
                chunks.push_back(std::string_view(listing).substr(offset, CURL_MAX_WRITE_SIZE));
            }
            std::vector<FilesInformationCurl::Entry> ordinary;
            const HRESULT ordinaryHr = ParseDirectoryListing(
                dos ? "01-02-26  03:04PM  4 f-04096.bin\r\n" : "-rw-r--r-- 1 owner group 4 Jan 02 2026 f-04096.bin\r\n", ordinary);
            for (const bool missing : {false, true})
            {
                std::vector<FilesInformationCurl::Entry> matches;
                uint64_t inspected   = 0u;
                uint64_t matchesSeen = 0u;
                bool correct         = SUCCEEDED(ordinaryHr) && ordinary.size() == 1u && ordinary.front().lastWriteTime != 0;
                uint64_t durationUs  = 0u;
                for (size_t attempt = 0u; attempt < repeats; ++attempt)
                {
                    uint64_t rows          = 0u;
                    const auto started     = std::chrono::steady_clock::now();
                    const HRESULT parseHr  = CurlDirectoryCursor::ParseChunksForSelfTest(chunks, matches, missing ? L"absent.bin" : L"f-04096.bin", &rows);
                    durationUs += Debug::Perf::ElapsedUs(started);
                    inspected += rows;
                    matchesSeen += matches.size();
                    correct = correct && parseHr == S_OK && rows == rowCount && matches.size() == (missing ? 0u : 1u);
                    if (! missing && matches.size() == 1u)
                    {
                        const auto& match = matches.front();
                        correct = correct && match.name == L"f-04096.bin" && match.sizeKnown && match.sizeBytes == 4u &&
                                  match.attributes == FILE_ATTRIBUTE_NORMAL && match.lastWriteTime == ordinary.front().lastWriteTime;
                    }
                }
                const std::wstring detail = std::format(L"{};queries={};missing={}", dos ? L"DOS" : L"Unix", repeats, missing);
                Debug::Perf::Emit(L"FileOps.Curl.EntryLookup.ParseOnly", detail.c_str(), durationUs, inspected, matchesSeen, correct ? S_OK : E_FAIL);
                DebugCheck(correct && inspected == repeats * rowCount, L"pure lookup parser validates every row and retains only exact matches", *passed, *failed);
            }
        }
        const std::vector<uint8_t> bytes{0x31u, 0x00u, 0xAAu, 0xFEu};
        constexpr std::string_view leaf = "literal%20 -> name.bin";
        const std::string record        = std::format("-rw-r--r-- 1 owner group 4 Jan 01 2026 {}\r\n", leaf);
        for (const Shape shape :
             {Shape::Literal, Shape::Missing, Shape::MalformedBefore, Shape::MalformedAfter, Shape::Duplicate, Shape::Outside, Shape::Overlong, Shape::Cancel})
        {
            for (const auto operation : {FILESYSTEM_COPY, FILESYSTEM_MOVE, FILESYSTEM_DELETE, FILESYSTEM_RENAME})
            {
                for (const bool bulk : {false, true})
                {
                    FakeFtpEndpoint endpoint(UploadRetention::Complete);
                    endpoint.SeedFile("/selected/" + std::string(shape == Shape::Missing ? "known.bin" : leaf), bytes);
                    endpoint.SeedFile("/sibling.bin", bytes);
                    endpoint.SeedDirectory("/destination");
                    if (shape == Shape::MalformedBefore || shape == Shape::MalformedAfter)
                    {
                        endpoint.SetDirectoryPayload("/selected",
                                                     shape == Shape::MalformedBefore ? "invalid LIST record\r\n" + record : record + "invalid LIST record\r\n");
                    }
                    else if (shape == Shape::Duplicate)
                    {
                        endpoint.SetDirectoryPayload("/selected", record + record);
                    }
                    else if (shape == Shape::Outside)
                    {
                        endpoint.SetDirectoryPayload("/selected", record + "-rw-r--r-- 1 owner group 4 Jan 01 2026 ../sibling.bin\r\n");
                    }
                    else if (shape == Shape::Overlong)
                    {
                        endpoint.SetDirectoryPayload("/selected", record + std::string(CURL_MAX_WRITE_SIZE + 1u, 'x') + "\r\n");
                    }
                    else if (shape == Shape::Cancel)
                    {
                        endpoint.SetLateListingGate("/selected", "/never-deleted", LateListingReply::WaitForCancel);
                    }
                    const EndpointSnapshot before = endpoint.Snapshot();
                    HRESULT hr                    = endpoint.Start();
                    wil::com_ptr<IFileSystem> fileSystem;
                    if (SUCCEEDED(hr))
                    {
                        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
                    }
                    FileSystemArenaOwner arena;
                    if (SUCCEEDED(hr))
                    {
                        hr = arena.Initialize(4096u);
                    }
                    if (! DebugCheck(SUCCEEDED(hr), L"entry lookup route fixture starts", *passed, *failed))
                    {
                        return E_FAIL;
                    }
                    const std::wstring prefix = std::format(L"//anonymous@127.0.0.1:{}", endpoint.Port());
                    const wchar_t* source     = CopyArenaString(arena.Get(), prefix + L"/selected/literal%20 -> name.bin");
                    const wchar_t* folder     = CopyArenaString(arena.Get(), prefix + L"/destination");
                    const wchar_t* target =
                        CopyArenaString(arena.Get(), prefix + (operation == FILESYSTEM_RENAME ? L"/selected/renamed.bin" : L"/destination/copied.bin"));
                    const wchar_t* newName = CopyArenaString(arena.Get(), L"renamed.bin");
                    auto* paths = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena.Get(), sizeof(const wchar_t*), alignof(const wchar_t*)));
                    auto* pair  = static_cast<FileSystemRenamePair*>(
                        AllocateFromFileSystemArena(arena.Get(), sizeof(FileSystemRenamePair), alignof(FileSystemRenamePair)));
                    if (! source || ! folder || ! target || ! newName || ! paths || ! pair)
                    {
                        DebugCheck(false, L"entry lookup arena is sufficient", *passed, *failed);
                        return E_FAIL;
                    }
                    paths[0] = source;
                    *pair    = FileSystemRenamePair{sizeof(FileSystemRenamePair), source, newName};
                    CancelControl control;
                    FileSystemOptions options{};
                    options.sizeBytes           = sizeof(options);
                    options.operationControl    = &control;
                    options.deadlineTickCount64 = GetTickCount64() + 15'000u;
                    CleanupDebtOperationCallback callback;
                    const auto started = std::chrono::steady_clock::now();
                    std::jthread worker([&]() noexcept
                    {
                        switch (operation)
                        {
                            case FILESYSTEM_COPY:
                                hr = bulk ? fileSystem->CopyItems(paths, 1u, folder, FILESYSTEM_FLAG_NONE, &options, &callback, nullptr)
                                          : fileSystem->CopyItem(source, target, FILESYSTEM_FLAG_NONE, &options, &callback, nullptr);
                                break;
                            case FILESYSTEM_MOVE:
                                hr = bulk ? fileSystem->MoveItems(paths, 1u, folder, FILESYSTEM_FLAG_NONE, &options, &callback, nullptr)
                                          : fileSystem->MoveItem(source, target, FILESYSTEM_FLAG_NONE, &options, &callback, nullptr);
                                break;
                            case FILESYSTEM_DELETE:
                                hr = bulk ? fileSystem->DeleteItems(paths, 1u, FILESYSTEM_FLAG_NONE, &options, &callback, nullptr)
                                          : fileSystem->DeleteItem(source, FILESYSTEM_FLAG_NONE, &options, &callback, nullptr);
                                break;
                            case FILESYSTEM_RENAME:
                                hr = bulk ? fileSystem->RenameItems(pair, 1u, FILESYSTEM_FLAG_NONE, &options, &callback, nullptr)
                                          : fileSystem->RenameItem(source, target, FILESYSTEM_FLAG_NONE, &options, &callback, nullptr);
                                break;
                            case FILESYSTEM_CREATE_DIRECTORY: hr = E_UNEXPECTED; break;
                        }
                    });
                    bool canceled     = true;
                    uint64_t cancelUs = 0u;
                    if (shape == Shape::Cancel)
                    {
                        const bool reached       = endpoint.WaitForLateListingReached(3s);
                        const auto cancelStarted = std::chrono::steady_clock::now();
                        control.abortRequested.store(true, std::memory_order_release);
                        const bool returned = WaitForSingleObject(worker.native_handle(), 2'500u) == WAIT_OBJECT_0;
                        cancelUs            = Debug::Perf::ElapsedUs(cancelStarted);
                        canceled            = reached && returned && cancelUs < 2'000'000u;
                        if (! returned)
                        {
                            endpoint.Stop(); // Emergency fixture unblock is never cooperative cancellation evidence.
                        }
                    }
                    worker.join();
                    const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
                    if (shape == Shape::Cancel && canceled)
                    {
                        // A canceled cursor must detach and clear borrowed callbacks before
                        // the same pool entry serves another operation on this endpoint.
                        control.abortRequested.store(false, std::memory_order_release);
                        ResolvedLocation recovered{};
                        FilesInformationCurl::Entry recoveredEntry{};
                        const HRESULT resolveHr =
                            ResolveLocation(Protocol::Ftp, FileSystemCurl::Settings{}, prefix + L"/sibling.bin", nullptr, true, recovered);
                        const CurlOperationOptionsScope recoveredScope(&options);
                        const HRESULT recoveredHr = FAILED(resolveHr) ? resolveHr : GetEntryInfo(recovered.connection, recovered.remotePath, recoveredEntry);
                        canceled                  = recoveredHr == S_OK && recoveredEntry.sizeKnown && recoveredEntry.sizeBytes == bytes.size();
                    }
                    fileSystem.reset();
                    endpoint.Stop();
                    const EndpointSnapshot after = endpoint.Snapshot();
                    const bool rejected          = shape != Shape::Literal;
                    const HRESULT expected       = shape == Shape::Literal    ? S_OK
                                                   : shape == Shape::Missing  ? HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)
                                                   : shape == Shape::Outside  ? HRESULT_FROM_WIN32(ERROR_INVALID_NAME)
                                                   : shape == Shape::Overlong ? HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW)
                                                   : shape == Shape::Cancel   ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                                              : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    const size_t mutations       = CountCommands(after, "STOR") + CountCommands(after, "MKD") + CountCommands(after, "DELE") +
                                                   CountCommands(after, "RMD") + CountCommands(after, "RNFR") + CountCommands(after, "RNTO");
                    const std::string finalPath  = operation == FILESYSTEM_RENAME ? "/selected/renamed.bin"
                                                   : bulk                         ? "/destination/" + std::string(leaf)
                                                                                  : "/destination/copied.bin";
                    const bool contents          = rejected ? before.files == after.files && before.directories == after.directories && mutations == 0u
                                                            : (operation == FILESYSTEM_COPY ? FileEquals(after, "/selected/" + std::string(leaf), bytes)
                                                                                            : FileAbsent(after, "/selected/" + std::string(leaf))) &&
                                                                  (operation == FILESYSTEM_DELETE || FileEquals(after, finalPath, bytes));
                    const bool correct = hr == expected && SUCCEEDED(after.serverHr) && contents && canceled && FileEquals(after, "/sibling.bin", bytes);
                    // Admission failed before a trustworthy source observation. A destructive
                    // receipt must not invent originalStillPresent, even with zero requests.
                    const bool truthful =
                        operation == FILESYSTEM_COPY ? callback.SingleCopyCompletionHasStatus(expected) : callback.SingleCompletionHasTruth(rejected);
                    const std::wstring detail = std::format(L"shape={};operation={};bulk={};hr=0x{:08X};expected=0x{:08X};contents={};truthful={}",
                                                            static_cast<unsigned int>(shape),
                                                            static_cast<unsigned int>(operation),
                                                            bulk,
                                                            static_cast<unsigned long>(hr),
                                                            static_cast<unsigned long>(expected),
                                                            contents,
                                                            truthful);
                    Debug::Perf::Emit(L"FileOps.Curl.EntryLookup.TruthFixture",
                                      detail.c_str(),
                                      durationUs,
                                      mutations,
                                      CountCommands(after, "LIST"),
                                      correct && truthful ? S_OK : E_FAIL);
                    if (shape == Shape::Cancel)
                    {
                        Debug::Perf::Emit(
                            L"FileOps.Curl.EntryLookup.ListCancel", detail.c_str(), cancelUs, mutations, canceled ? 1u : 0u, correct ? S_OK : E_FAIL);
                    }
                    DebugCheck(correct, std::format(L"{} validates the complete parent before mutation", detail).c_str(), *passed, *failed);
                    DebugCheck(truthful, L"entry lookup has one truthful terminal receipt", *passed, *failed);
                }
            }
        }
        for (const bool mixed : {false, true})
        {
            for (const bool dos : {false, true})
            {
                FakeFtpEndpoint endpoint(UploadRetention::Complete);
                endpoint.SeedDirectory("/wide");
                endpoint.SeedFile("/probe.bin", bytes);
                std::string listing;
                for (size_t index = 0u; index < 4097u; ++index)
                {
                    listing.append(dos ? std::format("01-02-26  03:04PM  4 f-{:05}.bin\r\n", index)
                                       : std::format("-rw-r--r-- 1 owner group 4 Jan 02 2026 f-{:05}.bin\r\n", index));
                }
                endpoint.SetDirectoryPayload("/wide", std::move(listing));
                HRESULT hr = endpoint.Start();
                ResolvedLocation resolved{};
                if (SUCCEEDED(hr))
                {
                    hr = ResolveLocation(Protocol::Ftp,
                                         FileSystemCurl::Settings{},
                                         std::format(L"//anonymous@127.0.0.1:{}/wide/f-04096.bin", endpoint.Port()),
                                         nullptr,
                                         true,
                                         resolved);
                }
                if (! DebugCheck(SUCCEEDED(hr), L"wide repeated entry lookup fixture starts", *passed, *failed))
                {
                    return E_FAIL;
                }
                CancelControl control;
                FileSystemOptions options{};
                options.sizeBytes           = sizeof(options);
                options.operationControl    = &control;
                options.deadlineTickCount64 = GetTickCount64() + 30'000u;
                const CurlOperationOptionsScope scope(&options);
                FilesInformationCurl::Entry entry{};
                std::vector<FilesInformationCurl::Entry> expectedMetadata;
                const HRESULT parseHr = ParseDirectoryListing(
                    dos ? "01-02-26  03:04PM  4 f-04096.bin\r\n" : "-rw-r--r-- 1 owner group 4 Jan 02 2026 f-04096.bin\r\n", expectedMetadata);
                CurlEntryLookupMetrics peak{};
                const auto started = std::chrono::steady_clock::now();
                bool correct       = SUCCEEDED(parseHr) && expectedMetadata.size() == 1u && expectedMetadata.front().lastWriteTime != 0;
                for (size_t attempt = 0u; attempt < 16u; ++attempt)
                {
                    CurlEntryLookupMetrics lookup{};
                    hr                 = GetEntryInfo(resolved.connection, resolved.remotePath, entry, &lookup);
                    correct            = correct && hr == S_OK && entry.name == L"f-04096.bin" && entry.sizeKnown && entry.sizeBytes == 4u &&
                                         entry.lastWriteTime == expectedMetadata.front().lastWriteTime && lookup.rows == 4097u &&
                                         lookup.pathBytes <= Common::FileOperations::kTraversalMaxQueuedPathBytes &&
                                         lookup.metadataBytes <= Common::FileOperations::kTraversalMaxMetadataBytes;
                    peak.pathBytes     = (std::max)(peak.pathBytes, lookup.pathBytes);
                    peak.metadataBytes = (std::max)(peak.metadataBytes, lookup.metadataBytes);
                    if (mixed)
                    {
                        uint64_t probeSize    = 0u;
                        bool probeKnown       = false;
                        const HRESULT probeHr = CurlProbeRemoteFileSize(resolved.connection, L"/probe.bin", probeSize, probeKnown);
                        correct               = correct && probeHr == S_OK && probeKnown && probeSize == bytes.size();
                    }
                }
                const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
                endpoint.Stop();
                const EndpointSnapshot after = endpoint.Snapshot();
                correct = correct && SUCCEEDED(after.serverHr) && CountCommands(after, "LIST") == 16u && after.passiveListenerCreations == 1u &&
                          after.passiveListenerReuses == 15u && CountCommands(after, "USER") == (mixed ? 2u : 1u) &&
                          (! mixed || CountCommands(after, "SIZE") == 16u) && after.dataTransfers[static_cast<size_t>(DataTransferKind::List)].firstUse == 1u &&
                          after.dataTransfers[static_cast<size_t>(DataTransferKind::List)].reuse == 15u;
                const std::wstring detail = std::format(L"{};queries=16;mixed={}", dos ? L"DOS" : L"Unix", mixed);
                Debug::Perf::Emit(
                    L"FileOps.Curl.EntryLookup.WideRepeated", detail.c_str(), durationUs, 4097u, CountCommands(after, "LIST"), correct ? S_OK : E_FAIL);
                DebugCheck(correct, L"repeated wide lookup returns complete requested metadata", *passed, *failed);
                Debug::Perf::Emit(L"FileOps.Curl.EntryLookup.Retained", detail.c_str(), 0u, peak.pathBytes, peak.metadataBytes, correct ? S_OK : E_FAIL);
                Debug::Perf::Emit(L"FileOps.Curl.EntryLookup.Connections",
                                  detail.c_str(),
                                  0u,
                                  after.passiveListenerCreations,
                                  after.passiveListenerReuses,
                                  correct ? S_OK : E_FAIL);
                Debug::Perf::Emit(L"FileOps.Curl.EntryLookup.ControlSessions",
                                  detail.c_str(),
                                  0u,
                                  CountCommands(after, "USER"),
                                  CountCommands(after, "SIZE"),
                                  correct ? S_OK : E_FAIL);
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: no fixture exception crosses into the host.
        Debug::Error(L"Curl entry lookup fixture failed after std::exception.");
        DebugCheck(false, L"entry lookup fixture must not throw", *passed, *failed);
    }
    return *failed == 0u ? S_OK : E_FAIL;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlCopyTraversalTruthForSelfTest(unsigned int* passed, unsigned int* failed) noexcept
{
    if (! passed || ! failed)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"Copy traversal fixture initializes Winsock", *passed, *failed))
    {
        return E_FAIL;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    try
    {
        {
            // Raw FTP client is intentionally fixture-local: it must leave an
            // unaccepted data connection behind a refused RETR, something a
            // production reader's earlier SIZE admission normally prevents.
            FakeFtpEndpoint endpoint(UploadRetention::Complete);
            const std::vector<uint8_t> expected{0x31u, 0x00u, 0xAAu, 0xFEu};
            endpoint.SeedFile("/known.bin", expected);
            const HRESULT started      = endpoint.Start();
            const auto connectLoopback = [](unsigned short port) noexcept
            {
                wil::unique_socket socketValue(socket(AF_INET, SOCK_STREAM, IPPROTO_TCP));
                sockaddr_in address{};
                address.sin_family      = AF_INET;
                address.sin_addr.s_addr = htonl(INADDR_LOOPBACK);
                address.sin_port        = htons(port);
                if (! socketValue || FAILED(ConfigureSocketTimeouts(socketValue.get())) ||
                    connect(socketValue.get(), reinterpret_cast<const sockaddr*>(&address), static_cast<int>(sizeof(address))) != 0)
                {
                    socketValue.reset();
                }
                return socketValue;
            };
            bool reusedPort      = false;
            const bool delivered = [&]()
            {
                if (FAILED(started))
                {
                    return false;
                }
                auto control = connectLoopback(endpoint.Port());
                if (! control)
                {
                    return false;
                }
                std::string pending;
                std::string reply;
                const auto readReply = [&](std::string_view prefix) noexcept
                { return SUCCEEDED(ReadControlLine(control.get(), {}, pending, reply)) && reply.starts_with(prefix); };
                const auto command = [&](std::string_view text, std::string_view prefix) noexcept
                { return SUCCEEDED(SendAll(control.get(), text)) && readReply(prefix); };
                const auto enterPassive = [&](unsigned short& port) noexcept
                {
                    if (! command("EPSV\r\n", "229 "))
                    {
                        return false;
                    }
                    const size_t begin = reply.find("|||");
                    if (begin == std::string::npos)
                    {
                        return false;
                    }
                    unsigned int parsedPort = 0u;
                    const auto parsed       = std::from_chars(reply.data() + begin + 3u, reply.data() + reply.size(), parsedPort);
                    if (parsed.ec != std::errc{} || parsed.ptr == reply.data() + reply.size() || *parsed.ptr != '|' || parsedPort == 0u ||
                        parsedPort > (std::numeric_limits<unsigned short>::max)())
                    {
                        return false;
                    }
                    port = static_cast<unsigned short>(parsedPort);
                    return true;
                };
                if (! readReply("220 ") || ! command("USER anonymous\r\n", "331 ") || ! command("PASS fixture\r\n", "230 "))
                {
                    return false;
                }
                unsigned short abandonedPort = 0u;
                if (! enterPassive(abandonedPort))
                {
                    return false;
                }
                auto abandoned = connectLoopback(abandonedPort);
                if (! abandoned || ! command("RETR /missing.bin\r\n", "550 "))
                {
                    return false;
                }
                unsigned short firstUsedPort = 0u;
                for (unsigned int attempt = 0u; attempt < 2u; ++attempt)
                {
                    unsigned short port = 0u;
                    if (! enterPassive(port))
                    {
                        return false;
                    }
                    if (attempt == 0u)
                    {
                        firstUsedPort = port;
                    }
                    else
                    {
                        reusedPort = port == firstUsedPort;
                    }
                    auto data = connectLoopback(port);
                    if (! data || ! command("RETR /known.bin\r\n", "150 "))
                    {
                        return false;
                    }
                    std::array<uint8_t, 8u> received{};
                    size_t total = 0u;
                    for (;;)
                    {
                        const int count = recv(data.get(), reinterpret_cast<char*>(received.data() + total), static_cast<int>(received.size() - total), 0);
                        if (count == 0)
                        {
                            break;
                        }
                        if (count < 0 || static_cast<size_t>(count) > expected.size() - total)
                        {
                            return false;
                        }
                        total += static_cast<size_t>(count);
                    }
                    if (total != expected.size() || ! std::equal(expected.begin(), expected.end(), received.begin()) || ! readReply("226 "))
                    {
                        return false;
                    }
                }
                return true;
            }();
            endpoint.Stop();
            const EndpointSnapshot after = endpoint.Snapshot();
            DebugCheck(delivered && SUCCEEDED(after.serverHr) && FileEquals(after, "/known.bin", expected),
                       L"a refused FTP RETR cannot leak its pending data connection into the next transfer",
                       *passed,
                       *failed);
            DebugCheck(reusedPort && after.passiveListenerCreations == 2u && after.passiveListenerReuses == 1u && after.passivePendingRetirements == 1u,
                       L"FTP fixture reuses consumed listeners and retires only the abandoned request",
                       *passed,
                       *failed);
            const DataTransferCounts& retrieveCounts = after.dataTransfers[static_cast<size_t>(DataTransferKind::Retrieve)];
            DebugCheck(retrieveCounts.firstUse == 1u && retrieveCounts.reuse == 1u,
                       L"FTP fixture counts accepted RETR first-use and reuse, not the refused transfer",
                       *passed,
                       *failed);
            Debug::Perf::Emit(L"FileOps.Curl.FakeFtp.PassiveReuseProof",
                              L"abandoned-RETR;two-complete-transfers",
                              0u,
                              after.passiveListenerCreations,
                              after.passiveListenerReuses,
                              delivered && reusedPort ? S_OK : E_FAIL);
        }
        enum class Shape
        {
            Shallow,
            Literal,
            Malformed,
            Outside,
            Overlong,
            Cancel,
            LateRefusal,
            LateCancel,
            Wide,
            Depth128,
            Depth129
        };
        const std::vector<uint8_t> bytes{0x31u, 0x00u, 0xAAu, 0xFEu};
        for (const Shape shape : {Shape::Shallow,
                                  Shape::Literal,
                                  Shape::Malformed,
                                  Shape::Outside,
                                  Shape::Overlong,
                                  Shape::Cancel,
                                  Shape::LateRefusal,
                                  Shape::LateCancel,
                                  Shape::Wide,
                                  Shape::Depth128,
                                  Shape::Depth129})
        {
            for (const bool move : {false, true})
            {
                if (move && shape != Shape::Depth128)
                {
                    continue; // The preflight gate covers the other Move failures and mode constraints.
                }
                for (const bool bulk : {false, true})
                {
                    if (bulk && shape == Shape::Wide)
                    {
                        continue; // Both queue modes stream 4,097 actual file publications; other shapes cover both APIs.
                    }
                    // Diagnose parallel transport before the long serial case can
                    // leave loopback resources under pressure. Keep the same corpus.
                    const std::array<unsigned int, 2u> concurrencyOrder{{shape == Shape::Wide ? 4u : 1u, shape == Shape::Wide ? 1u : 4u}};
                    for (const unsigned int concurrency : concurrencyOrder)
                    {
                        FakeFtpEndpoint source(UploadRetention::Complete);
                        FakeFtpEndpoint destination(UploadRetention::Complete);
                        std::string leaf       = "/selected/leaf.bin";
                        const bool late        = shape == Shape::LateRefusal || shape == Shape::LateCancel;
                        const bool cancel      = shape == Shape::Cancel || shape == Shape::LateCancel;
                        const bool earlyReject = shape == Shape::Malformed || shape == Shape::Outside || shape == Shape::Overlong || shape == Shape::Cancel;
                        const bool rejected    = earlyReject || late || shape == Shape::Depth129;
                        if (shape == Shape::Depth128 || shape == Shape::Depth129)
                        {
                            std::string directory = "/selected";
                            const size_t depth    = Common::FileOperations::kTraversalMaxDepth + (shape == Shape::Depth129 ? 1u : 0u);
                            for (size_t index = 0u; index < depth; ++index)
                            {
                                directory.append("/d");
                            }
                            leaf = directory + "/leaf.bin";
                        }
                        else if (shape == Shape::Literal)
                        {
                            leaf = "/selected/space %20 \"quoted\"/caf\xC3\xA9 ordinary -> suffix %25.bin";
                        }
                        else if (late)
                        {
                            source.SeedFile("/selected/a/early.bin", bytes);
                            leaf = "/selected/b/leaf.bin";
                        }
                        source.SeedFile(leaf, bytes);
                        source.SeedFile("/sibling.bin", bytes);
                        if (shape == Shape::Malformed)
                        {
                            source.SetDirectoryPayload("/selected", "not a supported LIST record\r\n");
                        }
                        else if (shape == Shape::Outside)
                        {
                            source.SetDirectoryPayload("/selected", "-rw-r--r-- 1 owner group 4 Jan 01 2026 ../sibling.bin\r\n");
                        }
                        else if (shape == Shape::Overlong)
                        {
                            source.SetDirectoryPayload("/selected", std::string(CURL_MAX_WRITE_SIZE + 1u, 'x') + "\r\n");
                        }
                        else if (shape == Shape::LateRefusal)
                        {
                            source.SetDirectoryAccessDenied("/selected/b");
                        }
                        else if (cancel)
                        {
                            source.SetLateListingGate(late ? "/selected/b" : "/selected", "/never-deleted", LateListingReply::WaitForCancel);
                        }
                        else if (shape == Shape::Wide)
                        {
                            for (size_t index = 0u; index < Common::FileOperations::kTraversalMaxQueuedEntries; ++index)
                            {
                                source.SeedFile(std::format("/selected/a-{:05}.bin", index), bytes);
                            }
                        }
                        const std::string destinationRoot = bulk ? "/destination/selected" : "/destination";
                        destination.SeedDirectory(destinationRoot);
                        destination.SeedFile("/sibling.bin", bytes);
                        const EndpointSnapshot sourceBefore      = source.Snapshot();
                        const EndpointSnapshot destinationBefore = destination.Snapshot();
                        HRESULT hr                               = source.Start();
                        if (SUCCEEDED(hr))
                        {
                            hr = destination.Start();
                        }
                        wil::com_ptr<IFileSystem> fileSystem;
                        if (SUCCEEDED(hr))
                        {
                            hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
                        }
                        wil::com_ptr<IInformations> information;
                        if (SUCCEEDED(hr))
                        {
                            hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
                        }
                        if (SUCCEEDED(hr))
                        {
                            hr = information->SetConfiguration(
                                std::format(
                                    R"({{"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":{},"deleteMaxConcurrency":1,"ftpUseEpsv":true}})",
                                    concurrency)
                                    .c_str());
                        }
                        FileSystemArenaOwner arena;
                        if (SUCCEEDED(hr))
                        {
                            hr = arena.Initialize(4096u);
                        }
                        if (! DebugCheck(SUCCEEDED(hr), L"Copy traversal fixture starts", *passed, *failed))
                        {
                            return E_FAIL;
                        }
                        const wchar_t* sourcePath      = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/selected", source.Port()));
                        const wchar_t* destinationPath = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/destination", destination.Port()));
                        auto* paths = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena.Get(), sizeof(const wchar_t*), alignof(const wchar_t*)));
                        if (! sourcePath || ! destinationPath || ! paths)
                        {
                            DebugCheck(false, L"Copy traversal caller arena is sufficient", *passed, *failed);
                            return E_FAIL;
                        }
                        paths[0] = sourcePath;
                        CancelControl control;
                        FileSystemOptions options{};
                        options.sizeBytes           = sizeof(options);
                        options.linkPolicy          = FILESYSTEM_LINK_PRESERVE;
                        options.operationControl    = &control;
                        // Serial 4,097-file loopback Copy was 108s on the archived
                        // Debug lane and now sits near 120s. Keep 180s of headroom
                        // so listing-before-descent and machine load cannot miss
                        // the last few publications on the documented 120s clock.
                        options.deadlineTickCount64 = GetTickCount64() + (shape == Shape::Wide ? 180'000u : 120'000u);
                        CleanupDebtOperationCallback callback;
                        constexpr FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
                        const std::wstring route        = std::format(L"shape={};{}{};concurrency={}",
                                                                      static_cast<unsigned int>(shape),
                                                                      move ? L"Move" : L"Copy",
                                                                      bulk ? L"Items" : L"Item",
                                                                      concurrency);
                        Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.Begin", route.c_str(), 0u, 0u, 0u, S_OK);
                        Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.EndpointPorts", route.c_str(), 0u, source.Port(), destination.Port(), S_OK);
                        const auto started = std::chrono::steady_clock::now();
                        std::jthread worker([&]() noexcept
                        {
                            hr = move ? (bulk ? fileSystem->MoveItems(paths, 1u, destinationPath, flags, &options, &callback, nullptr)
                                              : fileSystem->MoveItem(sourcePath, destinationPath, flags, &options, &callback, nullptr))
                                      : (bulk ? fileSystem->CopyItems(paths, 1u, destinationPath, flags, &options, &callback, nullptr)
                                              : fileSystem->CopyItem(sourcePath, destinationPath, flags, &options, &callback, nullptr));
                        });
                        bool cooperativeCancel = false;
                        uint64_t cancelUs      = 0u;
                        if (cancel)
                        {
                            const bool reached       = source.WaitForLateListingReached(5s);
                            const auto cancelStarted = std::chrono::steady_clock::now();
                            control.abortRequested.store(true, std::memory_order_release);
                            const bool returned = WaitForSingleObject(worker.native_handle(), 2'500u) == WAIT_OBJECT_0;
                            cancelUs            = Debug::Perf::ElapsedUs(cancelStarted);
                            cooperativeCancel   = reached && returned && cancelUs < 2'000'000u;
                            if (! returned)
                            {
                                source.Stop();
                            } // Owned emergency unblock is never a passing cancel.
                        }
                        worker.join();
                        const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
                        information.reset();
                        fileSystem.reset();
                        source.Stop();
                        destination.Stop();
                        const EndpointSnapshot sourceAfter      = source.Snapshot();
                        const EndpointSnapshot destinationAfter = destination.Snapshot();
                        Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.ControlSessions",
                                          route.c_str(),
                                          0u,
                                          CountCommands(sourceAfter, "USER"),
                                          CountCommands(destinationAfter, "USER"),
                                          S_OK);
                        const HRESULT expected        = shape == Shape::Malformed                              ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
                                                        : shape == Shape::Outside                              ? HRESULT_FROM_WIN32(ERROR_INVALID_NAME)
                                                        : shape == Shape::Overlong || shape == Shape::Depth129 ? HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW)
                                                        : shape == Shape::LateRefusal                          ? HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)
                                                        : cancel                                               ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                                                                               : S_OK;
                        auto expectedDestinationFiles = destinationBefore.files;
                        if (! rejected)
                        {
                            for (const auto& [path, contents] : sourceBefore.files)
                            {
                                if (path.starts_with("/selected/"))
                                {
                                    expectedDestinationFiles.insert_or_assign(destinationRoot + path.substr(9u), contents);
                                }
                            }
                        }
                        else if (late)
                        {
                            expectedDestinationFiles.insert_or_assign(destinationRoot + "/a/early.bin", bytes);
                        }
                        const bool sourceCorrect = move ? sourceAfter.files.size() == 1u && ! sourceAfter.directories.contains("/selected")
                                                        : sourceBefore.files == sourceAfter.files && sourceBefore.directories == sourceAfter.directories;
                        const bool contents      = sourceCorrect && destinationAfter.files == expectedDestinationFiles &&
                                                   (! earlyReject || destinationAfter.directories == destinationBefore.directories) &&
                                                   FileEquals(sourceAfter, "/sibling.bin", bytes) && FileEquals(destinationAfter, "/sibling.bin", bytes);
                        const bool truthful      = move ? callback.SingleCompletionHasTruth(false) : callback.SingleCopyCompletionHasStatus(hr);
                        bool boundedFixture = true;
                        std::wstring boundedDetail;
                        if (shape == Shape::Wide)
                        {
                            constexpr std::array<std::string_view, static_cast<size_t>(DataTransferKind::Count)> commands{{"LIST", "NLST", "RETR", "STOR"}};
                            for (const bool sourceEndpoint : {true, false})
                            {
                                const EndpointSnapshot& endpoint  = sourceEndpoint ? sourceAfter : destinationAfter;
                                uint64_t firstUses                = 0u;
                                uint64_t reuses                   = 0u;
                                bool commandCountsMatch           = true;
                                const std::wstring endpointDetail = std::format(L"{};endpoint={}", route, sourceEndpoint ? L"source" : L"destination");
                                for (size_t index = 0u; index < commands.size(); ++index)
                                {
                                    const DataTransferCounts& counts = endpoint.dataTransfers[index];
                                    firstUses += counts.firstUse;
                                    reuses += counts.reuse;
                                    const bool complete               = counts.firstUse + counts.reuse == CountCommands(endpoint, commands[index]);
                                    commandCountsMatch                = commandCountsMatch && complete;
                                    const std::wstring transferDetail = std::format(L"{};command={}", endpointDetail, Utf16FromUtf8(commands[index]));
                                    Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.DataTransferUses",
                                                      transferDetail.c_str(),
                                                      0u,
                                                      counts.firstUse,
                                                      counts.reuse,
                                                      complete ? S_OK : E_FAIL);
                                }
                                const size_t payloadIndex         = static_cast<size_t>(sourceEndpoint ? DataTransferKind::Retrieve : DataTransferKind::Store);
                                const DataTransferCounts& payload = endpoint.dataTransfers[payloadIndex];
                                // N successful transfers over K healthy sessions have N-K
                                // reuses, not necessarily N-1. Require complete accounting
                                // and the unchanged creation cap instead of one-session math.
                                const bool accounted = commandCountsMatch && payload.firstUse + payload.reuse == 4'097u &&
                                                       CountCommands(endpoint, "NLST") == 0u && endpoint.passivePendingRetirements == 0u &&
                                                       firstUses == endpoint.passiveListenerCreations && reuses == endpoint.passiveListenerReuses &&
                                                       endpoint.passiveListenerCreations <= 64u;
                                boundedFixture       = boundedFixture && accounted;
                                boundedDetail += std::format(L"[{};countsMatch={};payload={};nlst={};pendingRetirements={};firstUses={};creations={};reuses={};listenerReuses={}]",
                                                             sourceEndpoint ? L"source" : L"destination",
                                                             commandCountsMatch,
                                                             payload.firstUse + payload.reuse,
                                                             CountCommands(endpoint, "NLST"),
                                                             endpoint.passivePendingRetirements,
                                                             firstUses,
                                                             endpoint.passiveListenerCreations,
                                                             reuses,
                                                             endpoint.passiveListenerReuses);
                                Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.EndpointPassiveListeners",
                                                  endpointDetail.c_str(),
                                                  0u,
                                                  endpoint.passiveListenerCreations,
                                                  endpoint.passiveListenerReuses,
                                                  accounted ? S_OK : E_FAIL);
                                // Includes mutex wait and serialization, not transport time.
                                Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.ListPayloadBuild",
                                                  endpointDetail.c_str(),
                                                  endpoint.listPayloadBuildUs,
                                                  CountCommands(endpoint, "LIST"),
                                                  endpoint.listPayloadBytes,
                                                  S_OK);
                            }
                        }
                        const bool correct        = hr == expected && SUCCEEDED(sourceAfter.serverHr) && SUCCEEDED(destinationAfter.serverHr) && contents &&
                                                    (! cancel || cooperativeCancel) && boundedFixture;
                        const size_t publications = CountCommands(destinationAfter, "RNTO");
                        const std::wstring detail =
                            std::format(L"{};hr=0x{:08X};expected=0x{:08X};sourceServer=0x{:08X};destinationServer=0x{:08X};contents={};truthful={}",
                                        route,
                                        static_cast<unsigned long>(hr),
                                        static_cast<unsigned long>(expected),
                                        static_cast<unsigned long>(sourceAfter.serverHr),
                                        static_cast<unsigned long>(destinationAfter.serverHr),
                                        contents,
                                        truthful);
                        Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.TruthFixture",
                                          detail.c_str(),
                                          durationUs,
                                          sourceBefore.files.size() - 1u,
                                          publications,
                                          correct && truthful ? S_OK : E_FAIL);
                        if (shape == Shape::Wide)
                        {
                            Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.PassiveListeners",
                                              route.c_str(),
                                              0u,
                                              sourceAfter.passiveListenerCreations + destinationAfter.passiveListenerCreations,
                                              sourceAfter.passiveListenerReuses + destinationAfter.passiveListenerReuses,
                                              boundedFixture ? S_OK : E_FAIL);
                            Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.PassiveRetirements",
                                              route.c_str(),
                                              0u,
                                              sourceAfter.passivePendingRetirements + destinationAfter.passivePendingRetirements,
                                              sourceAfter.idleControlClosures + destinationAfter.idleControlClosures,
                                              S_OK);
                        }
                        if (cancel)
                        {
                            Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.ListCancel",
                                              route.c_str(),
                                              cancelUs,
                                              publications,
                                              cooperativeCancel ? 1u : 0u,
                                              correct ? S_OK : E_FAIL);
                        }
                        DebugCheck(correct,
                                   std::format(L"{} preserves exact contents, discovery failure and cooperative cancellation ({};cooperativeCancel={};cancelUs={};boundedFixture={}{})",
                                               route,
                                               detail,
                                               cooperativeCancel,
                                               cancelUs,
                                               boundedFixture,
                                               boundedDetail)
                                       .c_str(),
                                   *passed,
                                   *failed);
                        DebugCheck(truthful, std::format(L"{} reports one truthful completion", route).c_str(), *passed, *failed);
                    }
                }
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: report fixture failure without unwinding into the host.
        Debug::Error(L"Curl Copy traversal fixture failed after std::exception.");
        DebugCheck(false, L"Copy traversal fixture must not throw", *passed, *failed);
    }
    return *failed == 0u ? S_OK : E_FAIL;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlDeleteTraversalBoundsForSelfTest(unsigned int* passed, unsigned int* failed) noexcept
{
    if (! passed || ! failed)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"bounded traversal fixture initializes Winsock", *passed, *failed))
    {
        return E_FAIL;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    try
    {
        for (const Protocol protocol : {Protocol::Ftp, Protocol::Sftp, Protocol::Scp})
        {
            ConnectionInfo connection{};
            connection.protocol        = protocol;
            connection.basePath        = "/base%20space%25";
            connection.basePathWide    = L"/base space%";
            const bool ftp             = protocol == Protocol::Ftp;
            const std::string expected = ftp ? "/base space%/a \"quote\" caf\xC3\xA9.bin" : "\"/base space%/a \\\"quote\\\" caf\xC3\xA9.bin\"";
            DebugCheck(RemotePathForCommand(connection, L"/a \"quote\" caf\u00E9.bin") == expected,
                       L"native command paths use literal UTF-8 with protocol-specific argument grammar",
                       *passed,
                       *failed);
            DebugCheck(RemotePathForCommand(connection, L"/literal%20.bin") == (ftp ? "/base space%/literal%20.bin" : "\"/base space%/literal%20.bin\""),
                       L"native command paths do not decode a literal percent sequence",
                       *passed,
                       *failed);
            for (const std::wstring_view invalid :
                 {std::wstring_view(L"/bad\rname"), std::wstring_view(L"/bad\nname"), std::wstring_view(L"/bad\0name", 9u), std::wstring_view(L"/bad\xD800")})
            {
                DebugCheck(RemotePathForCommand(connection, invalid).empty(),
                           L"native command paths reject control injection, embedded NUL and invalid Unicode",
                           *passed,
                           *failed);
            }
        }
        enum class Shape
        {
            Wide,
            Deep,
            DepthLimit,
            LiteralArrow,
            Link,
            Malformed,
            Outside,
            Overlong,
            LiteralNested
        };
        const std::vector<uint8_t> bytes{0x34u, 0x00u, 0xFEu};
        for (const Shape shape : {Shape::Wide,
                                  Shape::Deep,
                                  Shape::DepthLimit,
                                  Shape::LiteralArrow,
                                  Shape::Link,
                                  Shape::Malformed,
                                  Shape::Outside,
                                  Shape::Overlong,
                                  Shape::LiteralNested})
        {
            for (const bool bulk : {false, true})
            {
                for (const unsigned int concurrency : {1u, 4u})
                {
                    FakeFtpEndpoint endpoint(UploadRetention::Complete);
                    size_t leafCount    = 1u;
                    const bool rejected = shape == Shape::DepthLimit || shape == Shape::Malformed || shape == Shape::Outside || shape == Shape::Overlong;
                    if (shape == Shape::Wide)
                    {
                        leafCount = Common::FileOperations::kTraversalMaxQueuedEntries + 1u;
                        for (size_t index = 0u; index < leafCount; ++index)
                        {
                            endpoint.SeedFile(std::format("/selected/leaf-{:05}.bin", index), bytes);
                        }
                    }
                    else if (shape == Shape::Deep || shape == Shape::DepthLimit)
                    {
                        std::string directory = "/selected";
                        const size_t depth    = shape == Shape::Deep ? 80u : Common::FileOperations::kTraversalMaxDepth + 1u;
                        for (size_t index = 0u; index < depth; ++index)
                        {
                            directory.append("/d");
                        }
                        endpoint.SeedFile(directory + "/leaf.bin", bytes);
                    }
                    else if (shape == Shape::LiteralNested)
                    {
                        endpoint.SeedFile("/selected/space %20 \"quoted\"/caf\xC3\xA9 %25.bin", bytes);
                    }
                    else
                    {
                        endpoint.SeedFile(shape == Shape::LiteralArrow ? "/selected/ordinary -> suffix.bin" : "/selected/leaf.bin", bytes);
                        if (shape == Shape::Link)
                        {
                            endpoint.SetDirectoryPayload("/selected", "lrwxrwxrwx 1 owner group 3 Jan 01 2026 leaf.bin -> /sibling.bin\r\n");
                        }
                        else if (shape == Shape::Malformed)
                        {
                            endpoint.SetDirectoryPayload("/selected", "not a supported LIST record\r\n");
                        }
                        else if (shape == Shape::Outside)
                        {
                            endpoint.SetDirectoryPayload("/selected", "-rw-r--r-- 1 owner group 3 Jan 01 2026 ../sibling.bin\r\n");
                        }
                        else if (shape == Shape::Overlong)
                        {
                            endpoint.SetDirectoryPayload("/selected", std::string(CURL_MAX_WRITE_SIZE + 1u, 'x') + "\r\n");
                        }
                    }
                    endpoint.SeedFile("/sibling.bin", bytes);
                    const EndpointSnapshot before = endpoint.Snapshot();
                    wil::com_ptr<IFileSystem> fileSystem;
                    HRESULT hr = endpoint.Start();
                    if (SUCCEEDED(hr))
                    {
                        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
                    }
                    wil::com_ptr<IInformations> information;
                    if (SUCCEEDED(hr))
                    {
                        hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
                    }
                    if (SUCCEEDED(hr))
                    {
                        hr = information->SetConfiguration(
                            std::format(
                                R"({{"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":1,"deleteMaxConcurrency":{},"ftpUseEpsv":true}})",
                                concurrency)
                                .c_str());
                    }
                    FileSystemArenaOwner arena;
                    if (SUCCEEDED(hr))
                    {
                        hr = arena.Initialize(4096u);
                    }
                    if (! DebugCheck(SUCCEEDED(hr), L"bounded native Delete fixture starts", *passed, *failed))
                    {
                        return E_FAIL;
                    }
                    const wchar_t* source = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/selected", endpoint.Port()));
                    auto* paths = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena.Get(), sizeof(const wchar_t*), alignof(const wchar_t*)));
                    if (! source || ! paths)
                    {
                        DebugCheck(false, L"bounded traversal caller arena is sufficient", *passed, *failed);
                        return E_FAIL;
                    }
                    paths[0] = source;
                    FileSystemOptions options{};
                    options.sizeBytes           = sizeof(options);
                    options.linkPolicy          = FILESYSTEM_LINK_PRESERVE;
                    options.deadlineTickCount64 = GetTickCount64() + 30'000u;
                    CleanupDebtOperationCallback callback;
                    const auto started        = std::chrono::steady_clock::now();
                    hr                        = bulk ? fileSystem->DeleteItems(paths, 1u, FILESYSTEM_FLAG_RECURSIVE, &options, &callback, nullptr)
                                                     : fileSystem->DeleteItem(source, FILESYSTEM_FLAG_RECURSIVE, &options, &callback, nullptr);
                    const uint64_t durationUs = Debug::Perf::ElapsedUs(started);
                    endpoint.Stop();
                    const EndpointSnapshot after = endpoint.Snapshot();
                    const HRESULT expected       = shape == Shape::Malformed ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
                                                   : shape == Shape::Outside ? HRESULT_FROM_WIN32(ERROR_INVALID_NAME)
                                                   : rejected                ? HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW)
                                                                             : S_OK;
                    const bool contents = rejected
                                              ? before.files == after.files && before.directories == after.directories &&
                                                    CountCommands(after, "DELE") + CountCommands(after, "RMD") == 0u
                                              : after.files.size() == 1u && ! after.directories.contains("/selected") && after.deletedPaths.size() == leafCount;
                    const bool correct  = hr == expected && SUCCEEDED(after.serverHr) && FileEquals(after, "/sibling.bin", bytes) && contents;
                    const bool truthful = callback.SingleCompletionHasTruth(rejected, FileSystemRouteContract::MutationClassification::RetryableNoCommit);
                    const std::wstring route =
                        std::format(L"shape={};{};concurrency={}", static_cast<unsigned int>(shape), bulk ? L"DeleteItems" : L"DeleteItem", concurrency);
                    const std::wstring detail = std::format(L"{};hr=0x{:08X};expected=0x{:08X};server=0x{:08X};contents={};truthful={};correct={}",
                                                            route,
                                                            static_cast<unsigned long>(hr),
                                                            static_cast<unsigned long>(expected),
                                                            static_cast<unsigned long>(after.serverHr),
                                                            contents,
                                                            truthful,
                                                            correct);
                    Debug::Perf::Emit(L"FileOps.Curl.NativeDelete.BoundsFixture",
                                      detail.c_str(),
                                      durationUs,
                                      leafCount,
                                      after.deletedPaths.size(),
                                      correct && truthful ? S_OK : E_FAIL);
                    DebugCheck(correct, std::format(L"{} has exact contents, status, and sibling preservation", route).c_str(), *passed, *failed);
                    DebugCheck(truthful, std::format(L"{} has one truthful terminal receipt", route).c_str(), *passed, *failed);
                }
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: report fixture failure to the host.
        Debug::Error(L"Curl Delete traversal bounds fixture failed after std::exception.");
        DebugCheck(false, L"bounded traversal fixture must not throw", *passed, *failed);
    }
    return *failed == 0u ? S_OK : E_FAIL;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlPartialTreeFailureForSelfTest(unsigned int* passed, unsigned int* failed) noexcept
{
    if (passed == nullptr || failed == nullptr)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"partial-tree fixture initializes Winsock", *passed, *failed))
    {
        return E_FAIL;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    try
    {
        const std::vector<uint8_t> bytes{0x31u, 0x00u, 0xAAu, 0xFEu};
        for (bool bulk : {false, true})
        {
            for (unsigned int concurrency : {1u, 4u})
            {
                FakeFtpEndpoint endpoint(UploadRetention::Complete);
                endpoint.SeedFile("/selected/a/removed.bin", bytes);
                endpoint.SeedFile("/selected/b/retained.bin", bytes);
                endpoint.SeedFile("/sibling.bin", bytes);
                endpoint.SetDirectoryAccessDenied("/selected/b");
                wil::com_ptr<IFileSystem> fileSystem;
                HRESULT hr = endpoint.Start();
                if (SUCCEEDED(hr))
                {
                    hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
                }
                wil::com_ptr<IInformations> information;
                if (SUCCEEDED(hr))
                {
                    hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
                }
                if (SUCCEEDED(hr))
                {
                    hr = information->SetConfiguration(
                        std::format(
                            R"({{"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":1,"deleteMaxConcurrency":{},"ftpUseEpsv":true}})",
                            concurrency)
                            .c_str());
                }
                FileSystemArenaOwner arena;
                if (SUCCEEDED(hr))
                {
                    hr = arena.Initialize(4096u);
                }
                if (! DebugCheck(SUCCEEDED(hr), L"partial-tree native Delete fixture starts", *passed, *failed))
                {
                    return E_FAIL;
                }
                const wchar_t* source = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/selected", endpoint.Port()));
                auto* paths           = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena.Get(), sizeof(const wchar_t*), alignof(const wchar_t*)));
                if (source == nullptr || paths == nullptr)
                {
                    DebugCheck(false, L"partial-tree fixture fits its bounded caller arena", *passed, *failed);
                    return E_FAIL;
                }
                paths[0] = source;
                CleanupDebtOperationCallback callback;
                const auto started = std::chrono::steady_clock::now();
                hr                 = bulk ? fileSystem->DeleteItems(paths, 1u, FILESYSTEM_FLAG_RECURSIVE, nullptr, &callback, nullptr)
                                          : fileSystem->DeleteItem(source, FILESYSTEM_FLAG_RECURSIVE, nullptr, &callback, nullptr);
                const uint64_t durationUs =
                    static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());
                endpoint.Stop(); // Drain peer-disconnect handling before checking server health.
                const EndpointSnapshot after = endpoint.Snapshot();
                const size_t mutations       = CountCommands(after, "DELE") + CountCommands(after, "RMD");
                const bool contents = after.files.size() == 2u && ! after.files.contains("/selected/a/removed.bin") &&
                                      ! after.directories.contains("/selected/a") && after.deletedPaths == std::vector<std::string>{"/selected/a/removed.bin"};
                const bool stateCorrect        = hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) && SUCCEEDED(after.serverHr) &&
                                                 after.rejectedDirectoryAccessCount != 0u && mutations == 2u && contents &&
                                                 FileEquals(after, "/selected/b/retained.bin", bytes) && FileEquals(after, "/sibling.bin", bytes) &&
                                                 after.directories.contains("/selected");
                const bool truthful            = callback.SingleCompletionHasTruth(true);
                const std::wstring route       = std::format(L"{};concurrency={}", bulk ? L"DeleteItems" : L"DeleteItem", concurrency);
                const std::wstring stateDetail = std::format(L"{};operationHr={:08X};serverHr={:08X};contents={};truthful={};files={};rootPresent={}",
                                                             route,
                                                             static_cast<unsigned long>(hr),
                                                             static_cast<unsigned long>(after.serverHr),
                                                             contents,
                                                             truthful,
                                                             after.files.size(),
                                                             after.directories.contains("/selected"));
                Debug::Perf::Emit(L"FileOps.Curl.PartialTreeFailure",
                                  stateDetail.c_str(),
                                  durationUs,
                                  mutations,
                                  after.rejectedDirectoryAccessCount,
                                  stateCorrect && truthful ? S_OK : E_FAIL);
                DebugCheck(stateCorrect, std::format(L"{} streams an earlier child before the later refusal", route).c_str(), *passed, *failed);
                DebugCheck(truthful, std::format(L"{} only claims retryable no-commit when no child mutation occurred", route).c_str(), *passed, *failed);
            }
        }

        // A different selected item can succeed without contaminating the refused item's
        // proof, whether the two top-level items execute serially or on separate workers.
        for (const unsigned int concurrency : {1u, 4u})
        {
            FakeFtpEndpoint endpoint(UploadRetention::Complete);
            endpoint.SeedFile("/complete/leaf.bin", bytes);
            endpoint.SeedFile("/selected/b/retained.bin", bytes);
            endpoint.SeedFile("/sibling.bin", bytes);
            // Refuse this selected root's first listing: the other selected item's
            // mutations must not contaminate this untouched item's no-commit proof.
            endpoint.SetDirectoryAccessDenied("/selected");
            HRESULT hr = endpoint.Start();
            wil::com_ptr<IFileSystem> fileSystem;
            if (SUCCEEDED(hr))
            {
                hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
            }
            wil::com_ptr<IInformations> information;
            if (SUCCEEDED(hr))
            {
                hr = fileSystem->QueryInterface(__uuidof(IInformations), information.put_void());
            }
            if (SUCCEEDED(hr))
            {
                const std::string config = std::format(
                    R"({{"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":1,"deleteMaxConcurrency":{},"ftpUseEpsv":true}})",
                    concurrency);
                hr = information->SetConfiguration(config.c_str());
            }
            FileSystemArenaOwner arena;
            if (SUCCEEDED(hr))
            {
                hr = arena.Initialize(4096u);
            }
            if (! DebugCheck(SUCCEEDED(hr), L"independent selected-item proof fixture starts", *passed, *failed))
            {
                return E_FAIL;
            }
            auto* paths = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena.Get(), 2u * sizeof(const wchar_t*), alignof(const wchar_t*)));
            const wchar_t* complete = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/complete", endpoint.Port()));
            const wchar_t* selected = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/selected", endpoint.Port()));
            if (paths == nullptr || complete == nullptr || selected == nullptr)
            {
                DebugCheck(false, L"independent selected-item fixture fits its caller arena", *passed, *failed);
                return E_FAIL;
            }
            paths[0] = complete;
            paths[1] = selected;
            CleanupDebtOperationCallback callback;
            const auto started = std::chrono::steady_clock::now();
            hr                 = fileSystem->DeleteItems(
                paths, 2u, static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, &callback, nullptr);
            const uint64_t durationUs =
                static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());
            const EndpointSnapshot after = endpoint.Snapshot();
            const size_t mutations       = CountCommands(after, "DELE") + CountCommands(after, "RMD");
            const bool stateCorrect      = FAILED(hr) && SUCCEEDED(after.serverHr) && mutations == 2u && after.rejectedDirectoryAccessCount == 1u &&
                                           after.files.size() == 2u && FileEquals(after, "/selected/b/retained.bin", bytes) &&
                                           FileEquals(after, "/sibling.bin", bytes) && ! after.directories.contains("/complete") &&
                                           after.directories.contains("/selected") && after.deletedPaths == std::vector<std::string>{"/complete/leaf.bin"};
            const bool truthful          = callback.TwoCompletionsKeepIndependentTruth();
            const std::wstring route     = std::format(L"DeleteItems;concurrency={}", concurrency);
            Debug::Perf::Emit(L"FileOps.Curl.IndependentItemProof",
                              route.c_str(),
                              durationUs,
                              mutations,
                              after.rejectedDirectoryAccessCount,
                              stateCorrect && truthful ? S_OK : E_FAIL);
            DebugCheck(stateCorrect, L"independent selected-item proof fixture preserves the refused tree and sibling", *passed, *failed);
            DebugCheck(truthful, L"another selected item's mutations never invalidate this item's no-mutation proof", *passed, *failed);
        }

        // The shared completion owner must never manufacture no-commit from an error code.
        // Explicit pre-mutation evidence remains usable for every destructive operation.
        for (const auto operation : {FILESYSTEM_DELETE, FILESYSTEM_MOVE, FILESYSTEM_RENAME})
        {
            for (const bool provedNoMutation : {false, true})
            {
                CleanupDebtOperationCallback callback;
                FileOperationProgress progress;
                const HRESULT initializeHr = progress.Initialize(operation, 1u, nullptr, &callback, nullptr);
                const HRESULT completionHr =
                    SUCCEEDED(initializeHr) ? progress.ReportItemCompleted(0u, L"/source", {}, E_ACCESSDENIED, provedNoMutation) : initializeHr;
                DebugCheck(SUCCEEDED(completionHr) &&
                               callback.SingleCompletionHasTruth(true,
                                                                 provedNoMutation ? FileSystemRouteContract::MutationClassification::RetryableNoCommit
                                                                                  : FileSystemRouteContract::MutationClassification::Indeterminate),
                           L"destructive completion requires explicit per-item no-mutation evidence",
                           *passed,
                           *failed);
            }
        }

        // Ordinary collisions are refused before even the rollback rename. Removing the
        // HRESULT guess must not turn these proved no-mutation paths into unknown outcomes.
        for (const auto operation : {FILESYSTEM_MOVE, FILESYSTEM_RENAME})
        {
            for (const bool bulk : {false, true})
            {
                FakeFtpEndpoint endpoint(UploadRetention::Complete);
                const bool move               = operation == FILESYSTEM_MOVE;
                const std::string destination = move ? "/occupied/source.bin" : "/selected/taken.bin";
                endpoint.SeedFile("/selected/source.bin", bytes);
                endpoint.SeedFile(destination, {0x90u, 0x91u});
                const EndpointSnapshot before = endpoint.Snapshot();
                HRESULT hr                    = endpoint.Start();
                wil::com_ptr<IFileSystem> fileSystem;
                if (SUCCEEDED(hr))
                {
                    hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
                }
                FileSystemArenaOwner arena;
                if (SUCCEEDED(hr))
                {
                    hr = arena.Initialize(4096u);
                }
                if (! DebugCheck(SUCCEEDED(hr), L"native collision no-commit fixture starts", *passed, *failed))
                {
                    return E_FAIL;
                }
                const std::wstring prefix = std::format(L"//anonymous@127.0.0.1:{}", endpoint.Port());
                const wchar_t* source     = CopyArenaString(arena.Get(), prefix + L"/selected/source.bin");
                const wchar_t* target     = CopyArenaString(arena.Get(), prefix + (move ? L"/occupied/source.bin" : L"/selected/taken.bin"));
                const wchar_t* folder     = CopyArenaString(arena.Get(), prefix + L"/occupied");
                const wchar_t* newName    = CopyArenaString(arena.Get(), L"taken.bin");
                auto* paths = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena.Get(), sizeof(const wchar_t*), alignof(const wchar_t*)));
                auto* pair =
                    static_cast<FileSystemRenamePair*>(AllocateFromFileSystemArena(arena.Get(), sizeof(FileSystemRenamePair), alignof(FileSystemRenamePair)));
                if (source == nullptr || target == nullptr || folder == nullptr || newName == nullptr || paths == nullptr || pair == nullptr)
                {
                    DebugCheck(false, L"native collision fixture fits its caller arena", *passed, *failed);
                    return E_FAIL;
                }
                paths[0]         = source;
                pair->sizeBytes  = sizeof(FileSystemRenamePair);
                pair->sourcePath = source;
                pair->newName    = newName;
                CleanupDebtOperationCallback callback;
                const auto started = std::chrono::steady_clock::now();
                if (move)
                {
                    hr = bulk ? fileSystem->MoveItems(paths, 1u, folder, FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr)
                              : fileSystem->MoveItem(source, target, FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
                }
                else
                {
                    hr = bulk ? fileSystem->RenameItems(pair, 1u, FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr)
                              : fileSystem->RenameItem(source, target, FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
                }
                const uint64_t durationUs =
                    static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());
                const EndpointSnapshot after = endpoint.Snapshot();
                const size_t mutations       = CountCommands(after, "RNFR") + CountCommands(after, "RNTO") + CountCommands(after, "DELE") +
                                               CountCommands(after, "RMD") + CountCommands(after, "MKD");
                const bool stateCorrect      = hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) && SUCCEEDED(after.serverHr) && mutations == 0u &&
                                               after.files == before.files && after.directories == before.directories;
                const bool truthful          = callback.SingleCompletionHasTruth(true, FileSystemRouteContract::MutationClassification::RetryableNoCommit);
                const std::wstring route     = std::format(L"{}{}", move ? L"Move" : L"Rename", bulk ? L"Items" : L"Item");
                Debug::Perf::Emit(L"FileOps.Curl.CollisionNoCommit", route.c_str(), durationUs, mutations, 0u, stateCorrect && truthful ? S_OK : E_FAIL);
                DebugCheck(stateCorrect, L"native collision preserves both source and occupant without a mutation request", *passed, *failed);
                DebugCheck(truthful, L"native collision retains explicit safe-retry proof", *passed, *failed);
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: no fixture exception crosses into the host.
        Debug::Error(L"Curl partial-tree fixture failed after std::exception.");
        DebugCheck(false, L"partial-tree fixture must not throw", *passed, *failed);
    }
    return *failed == 0u ? S_OK : E_FAIL;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlLostMutationReplyForSelfTest(unsigned int* passed, unsigned int* failed) noexcept
{
    if (passed == nullptr || failed == nullptr)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"lost mutation reply fixture initializes Winsock", *passed, *failed))
    {
        return E_FAIL;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    try
    {
        const std::vector<uint8_t> original{0x31u, 0x00u, 0x7Fu, 0xAAu};
        const std::vector<uint8_t> replacement{0xC5u, 0xFEu, 0x10u};
        for (bool rename : {false, true})
        {
            for (bool bulk : {false, true})
            {
                for (bool responseLost : {false, true})
                {
                    FakeFtpEndpoint endpoint(UploadRetention::Complete);
                    endpoint.SeedFile("/source.bin", original);
                    endpoint.SeedFile("/sibling.bin", original);
                    if (responseLost)
                    {
                        endpoint.SetLostMutationReply(rename ? "RNTO" : "DELE", "/source.bin", replacement);
                    }
                    wil::com_ptr<IFileSystem> fileSystem;
                    HRESULT hr = endpoint.Start();
                    if (SUCCEEDED(hr))
                    {
                        hr = CreateConfiguredFtpSelfTestFileSystem(fileSystem);
                    }
                    FileSystemArenaOwner arena;
                    if (SUCCEEDED(hr))
                    {
                        hr = arena.Initialize(4096u);
                    }
                    if (! DebugCheck(SUCCEEDED(hr), L"lost mutation reply fixture starts", *passed, *failed))
                    {
                        return E_FAIL;
                    }
                    const wchar_t* source      = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/source.bin", endpoint.Port()));
                    const wchar_t* destination = CopyArenaString(arena.Get(), std::format(L"//anonymous@127.0.0.1:{}/destination.bin", endpoint.Port()));
                    const wchar_t* leaf        = CopyArenaString(arena.Get(), L"destination.bin");
                    auto* paths = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena.Get(), sizeof(const wchar_t*), alignof(const wchar_t*)));
                    auto* pair  = static_cast<FileSystemRenamePair*>(
                        AllocateFromFileSystemArena(arena.Get(), sizeof(FileSystemRenamePair), alignof(FileSystemRenamePair)));
                    if (source == nullptr || destination == nullptr || leaf == nullptr || paths == nullptr || pair == nullptr)
                    {
                        DebugCheck(false, L"lost mutation reply fixture fits its bounded caller arena", *passed, *failed);
                        return E_FAIL;
                    }
                    paths[0] = source;
                    *pair    = FileSystemRenamePair{sizeof(FileSystemRenamePair), source, leaf};
                    CleanupDebtOperationCallback callback;
                    const auto started = std::chrono::steady_clock::now();
                    if (rename)
                    {
                        hr = bulk ? fileSystem->RenameItems(pair, 1u, FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr)
                                  : fileSystem->RenameItem(source, destination, FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
                    }
                    else
                    {
                        hr = bulk ? fileSystem->DeleteItems(paths, 1u, FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr)
                                  : fileSystem->DeleteItem(source, FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
                    }
                    const uint64_t durationUs =
                        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());
                    const EndpointSnapshot after = endpoint.Snapshot();
                    const size_t requests        = CountCommands(after, rename ? "RNTO" : "DELE");
                    const bool originalApplied =
                        rename ? FileEquals(after, "/destination.bin", original) : after.deletedPaths == std::vector<std::string>{"/source.bin"};
                    const bool sourceCorrect   = responseLost ? FileEquals(after, "/source.bin", replacement) : ! after.files.contains("/source.bin");
                    const size_t expectedFiles = 1u + (rename ? 1u : 0u) + (responseLost ? 1u : 0u);
                    const bool stateCorrect    = SUCCEEDED(after.serverHr) && after.files.size() == expectedFiles && originalApplied && sourceCorrect &&
                                                 FileEquals(after, "/sibling.bin", original) && after.directories == std::set<std::string>{"/"} &&
                                                 after.lostMutationReplyCount == (responseLost ? 1u : 0u);
                    const bool once            = requests == 1u && (! rename || CountCommands(after, "RNFR") == 1u) && CountCommands(after, "NLST") == 0u;
                    const bool truth           = (responseLost ? FAILED(hr) : hr == S_OK) && callback.SingleCompletionHasTruth(responseLost);
                    const std::wstring route   = std::format(L"{}{};responseLost={}", rename ? L"Rename" : L"Delete", bulk ? L"Items" : L"Item", responseLost);
                    Debug::Perf::Emit(L"FileOps.Curl.MutationReplyLoss",
                                      route.c_str(),
                                      durationUs,
                                      requests,
                                      after.lostMutationReplyCount,
                                      stateCorrect && once && truth ? S_OK : E_FAIL);
                    DebugCheck(stateCorrect, std::format(L"{} preserves the committed effect, replacement and sibling bytes", route).c_str(), *passed, *failed);
                    DebugCheck(once, std::format(L"{} sends exactly one mutation, never a replay", route).c_str(), *passed, *failed);
                    DebugCheck(
                        truth, std::format(L"{} reports exactly one honest completion, without safe-Retry proof after loss", route).c_str(), *passed, *failed);
                }
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: fail the test without escaping into the host.
        Debug::Error(L"Curl lost mutation reply fixture failed after std::exception.");
        DebugCheck(false, L"lost mutation reply fixture must not throw", *passed, *failed);
    }
    return *failed == 0u ? S_OK : E_FAIL;
}

// Host self-tests drive the deterministic fake FTP fixture through these exports: start one on a
// loopback port, run File Operations against `//anonymous@127.0.0.1:<port>/...`, then stop it.
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlStartFakeFtpForSelfTest(unsigned int* port, void** endpoint) noexcept
{
    if (port == nullptr || endpoint == nullptr)
    {
        return E_POINTER;
    }
    *port     = 0u;
    *endpoint = nullptr;
    WSADATA winsockData{};
    if (WSAStartup(MAKEWORD(2, 2), &winsockData) != 0)
    {
        return E_FAIL;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    try
    {
        auto fixture     = std::make_unique<FakeFtpEndpoint>(UploadRetention::Complete);
        const HRESULT hr = fixture->Start();
        if (FAILED(hr))
        {
            return hr;
        }
        *port     = fixture->Port();
        *endpoint = fixture.release();
        cleanupWinsock.release(); // Paired with the exported stop call while the DLL is pinned.
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary; owned sockets/Winsock unwind before returning failure.
        Debug::Error(L"Curl fake FTP start failed after std::exception.");
        return E_FAIL;
    }
}

extern "C" __declspec(dllexport) void __stdcall RedSalamanderCurlStopFakeFtpForSelfTest(void* endpoint) noexcept
{
    if (endpoint == nullptr)
    {
        return;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    std::unique_ptr<FakeFtpEndpoint> fixture(static_cast<FakeFtpEndpoint*>(endpoint));
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlDirectoryMoveForSelfTest(
    void* endpoint, BOOL arm, BOOL refuse, unsigned int* renameFromCount, unsigned int* renameToCount, BOOL* stateCorrect) noexcept
{
    if (! endpoint || ! renameFromCount || ! renameToCount || ! stateCorrect)
    {
        return E_POINTER;
    }
    *renameFromCount = 0u;
    *renameToCount = 0u;
    *stateCorrect = FALSE;
    try
    {
        auto& fixture = *static_cast<FakeFtpEndpoint*>(endpoint);
        const std::vector<uint8_t> bytes{0x31u, 0x00u, 0x7Fu, 0xAAu};
        if (arm != FALSE)
        {
            fixture.SeedFile("/br5/source/a.bin", bytes);
            fixture.SeedFile("/br5/source/nested/deep.bin", bytes);
            fixture.SeedDirectory("/br5/source/empty");
            fixture.SeedDirectory("/br5/destination");
            fixture.SeedFile("/br5/source-peer/keep.bin", bytes);
            fixture.SeedFile("/br5/sibling.bin", bytes);
            if (refuse != FALSE)
            {
                fixture.SetDirectoryRenameRefused("/br5/source");
            }
            return S_OK;
        }
        const EndpointSnapshot snapshot = fixture.Snapshot();
        *renameFromCount = static_cast<unsigned int>(CountCommands(snapshot, "RNFR"));
        *renameToCount = static_cast<unsigned int>(CountCommands(snapshot, "RNTO"));
        const std::string retainedRoot = refuse != FALSE ? "/br5/source" : "/br5/destination/source";
        const std::string absentRoot = refuse != FALSE ? "/br5/destination/source" : "/br5/source";
        const bool noRelayOrCleanup = CountCommands(snapshot, "RETR") == 0u && CountCommands(snapshot, "STOR") == 0u &&
                                      CountCommands(snapshot, "DELE") == 0u && CountCommands(snapshot, "RMD") == 0u && CountCommands(snapshot, "MKD") == 0u;
        *stateCorrect = SUCCEEDED(snapshot.serverHr) && noRelayOrCleanup && snapshot.files.size() == 4u &&
                        FileEquals(snapshot, retainedRoot + "/a.bin", bytes) && FileEquals(snapshot, retainedRoot + "/nested/deep.bin", bytes) &&
                        snapshot.directories.contains(retainedRoot + "/empty") && snapshot.directories.contains(retainedRoot + "/nested") &&
                        ! snapshot.directories.contains(absentRoot) && ! snapshot.files.contains(absentRoot + "/a.bin") &&
                        ! snapshot.files.contains(absentRoot + "/nested/deep.bin") && FileEquals(snapshot, "/br5/source-peer/keep.bin", bytes) &&
                        FileEquals(snapshot, "/br5/sibling.bin", bytes) && snapshot.deletedPaths.empty() ? TRUE : FALSE;
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: fixture exceptions cannot cross into the host.
        Debug::Error(L"Curl directory Move fixture failed after std::exception.");
        return E_FAIL;
    }
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlCommittedDeleteForSelfTest(
    void* endpoint, BOOL arm, unsigned int* requests, unsigned int* commits, BOOL* stateCorrect) noexcept
{
    if (endpoint == nullptr || requests == nullptr || commits == nullptr || stateCorrect == nullptr)
    {
        return E_POINTER;
    }
    *requests     = 0u;
    *commits      = 0u;
    *stateCorrect = FALSE;
    try
    {
        auto& fixture = *static_cast<FakeFtpEndpoint*>(endpoint);
        const std::vector<uint8_t> original{0x31u, 0x00u, 0x7Fu, 0xAAu};
        const std::vector<uint8_t> replacement{0xC5u, 0xFEu, 0x10u};
        if (arm != FALSE)
        {
            fixture.SeedFile("/c0/source.bin", original);
            fixture.SeedFile("/c0/sibling.bin", original);
            fixture.SetLostMutationReply("DELE", "/c0/source.bin", replacement);
            return S_OK;
        }
        const EndpointSnapshot snapshot = fixture.Snapshot();
        *requests                       = static_cast<unsigned int>(CountCommands(snapshot, "DELE", "/c0/source.bin"));
        *commits                        = static_cast<unsigned int>(snapshot.lostMutationReplyCount);
        *stateCorrect = SUCCEEDED(snapshot.serverHr) && snapshot.files.size() == 2u && FileEquals(snapshot, "/c0/source.bin", replacement) &&
                                FileEquals(snapshot, "/c0/sibling.bin", original) && snapshot.deletedPaths == std::vector<std::string>{"/c0/source.bin"}
                            ? TRUE
                            : FALSE;
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: no fixture exception crosses into the host.
        Debug::Error(L"Curl committed Delete fixture failed after std::exception.");
        return E_FAIL;
    }
}
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlPartialDeleteForSelfTest(
    void* endpoint, BOOL arm, unsigned int* requests, unsigned int* commits, BOOL* stateCorrect) noexcept
{
    if (endpoint == nullptr || requests == nullptr || commits == nullptr || stateCorrect == nullptr)
    {
        return E_POINTER;
    }
    *requests = 0u;
    *commits      = 0u;
    *stateCorrect = FALSE;
    try
    {
        auto& fixture = *static_cast<FakeFtpEndpoint*>(endpoint);
        const std::vector<uint8_t> bytes{0x31u, 0x00u, 0xAAu, 0xFEu};
        if (arm != FALSE)
        {
            fixture.SeedFile("/selected/a/removed.bin", bytes);
            fixture.SeedFile("/selected/b/retained.bin", bytes);
            fixture.SeedFile("/sibling.bin", bytes);
            fixture.SetDirectoryAccessDenied("/selected/b");
            return S_OK;
        }
        const EndpointSnapshot snapshot = fixture.Snapshot();
        *requests                       = static_cast<unsigned int>(CountCommands(snapshot, "DELE") + CountCommands(snapshot, "RMD"));
        *commits                        = static_cast<unsigned int>(snapshot.deletedPaths.size());
        *stateCorrect                   = SUCCEEDED(snapshot.serverHr) && snapshot.rejectedDirectoryAccessCount == 1u && snapshot.files.size() == 2u &&
                                                  FileEquals(snapshot, "/selected/b/retained.bin", bytes) && FileEquals(snapshot, "/sibling.bin", bytes) &&
                                                  snapshot.directories.contains("/selected") && ! snapshot.directories.contains("/selected/a") &&
                                                  snapshot.deletedPaths == std::vector<std::string>{"/selected/a/removed.bin"}
                                              ? TRUE
                                              : FALSE;
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: no fixture exception crosses into the host.
        Debug::Error(L"Curl partial Delete fixture failed after std::exception.");
        return E_FAIL;
    }
}

// C0 IMAP transport witness: the production listing/lookup owners run against a loopback
// IMAP server through the real libcurl transport. This is the only coverage that can
// observe libcurl's channel routing for listing-shaped versus bare FETCH sets; the
// synchronous ImapListingFixture bypasses libcurl entirely. Two checks are libcurl
// contract controls: if a newer libcurl starts writing listing rows as body, they fail
// and the capture policy in ImapResponseCapture must be re-evaluated, not silently kept.
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlImapTransportTruthForSelfTest(unsigned int* passed, unsigned int* failed) noexcept
{
    if (passed == nullptr || failed == nullptr)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"IMAP transport fixture initializes Winsock", *passed, *failed))
    {
        return E_FAIL;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });
    try
    {
        // ASCII on purpose: the literal *shape* is what changes libcurl's routing, and the
        // leaf-name assertion must not depend on the header charset fallback policy.
        const std::string literalSubjectUtf8 = "Literal subject delivered in braces";
        const std::wstring literalSubject    = L"Literal subject delivered in braces";
        FakeImapEndpoint server;
        server.AddMessage(FakeImapMessage{.uid = 1u, .subjectUtf8 = "Plain first", .literalSubject = false, .sizeBytes = 100u});
        server.AddMessage(FakeImapMessage{.uid = 2u, .subjectUtf8 = literalSubjectUtf8, .literalSubject = true, .sizeBytes = 200u});
        server.AddMessage(FakeImapMessage{.uid = 3u, .subjectUtf8 = "Plain third", .literalSubject = false, .sizeBytes = 300u});
        HRESULT hr = server.Start();
        if (SUCCEEDED(hr))
        {
            hr = EnsureCurlInitialized();
        }
        if (! DebugCheck(SUCCEEDED(hr), L"IMAP transport fixture starts its loopback server", *passed, *failed))
        {
            return E_FAIL;
        }

        ConnectionInfo conn{};
        conn.protocol           = Protocol::Imap;
        conn.host               = "127.0.0.1";
        conn.port               = server.Port();
        conn.user               = "selftest";
        conn.password           = "selftest-secret";
        conn.basePath           = "/";
        conn.basePathWide       = L"/";
        conn.limiterKey         = std::format(L"imap-transport-selftest:{}", server.Port());
        conn.connectTimeoutMs   = 5000u;
        conn.operationTimeoutMs = 20000u;
        FileSystemOptions options{};
        options.sizeBytes = sizeof(options);
        const CurlOperationOptionsScope operationScope(&options);

        // 1. Listing: one bulk listing-shaped FETCH must deliver every row, including the
        //    literal subject, without any repair request.
        std::vector<FilesInformationCurl::Entry> entries;
        const auto listingStarted = std::chrono::steady_clock::now();
        hr                        = ReadDirectoryEntries(conn, L"/INBOX", entries);
        const uint64_t listingUs  = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - listingStarted).count());
        ImapEndpointSnapshot snapshot = server.Snapshot();
        const auto findEntry          = [&entries](uint64_t uid) noexcept -> const FilesInformationCurl::Entry*
        {
            const auto found = std::ranges::find_if(entries, [uid](const auto& entry) noexcept { return entry.fileIndex == uid; });
            return found == entries.end() ? nullptr : &*found;
        };
        const FilesInformationCurl::Entry* first  = findEntry(1u);
        const FilesInformationCurl::Entry* second = findEntry(2u);
        const FilesInformationCurl::Entry* third  = findEntry(3u);
        DebugCheck(hr == S_OK && entries.size() == 3u && first != nullptr && second != nullptr && third != nullptr,
                   L"real-libcurl IMAP listing publishes every message from the bulk FETCH",
                   *passed,
                   *failed);
        DebugCheck(first != nullptr && second != nullptr && third != nullptr && first->sizeKnown && first->sizeBytes == 100u && second->sizeKnown &&
                       second->sizeBytes == 200u && third->sizeKnown && third->sizeBytes == 300u,
                   L"listing sizes come from RFC822.SIZE, not from a 0 B fallback",
                   *passed,
                   *failed);
        DebugCheck(second != nullptr && second->name == BuildImapMessageLeafName(literalSubject, {}, 777u, 2u),
                   L"a literal ENVELOPE subject survives libcurl line delivery and decodes",
                   *passed,
                   *failed);
        DebugCheck(snapshot.fetchSets.size() == 1u && snapshot.fetchSets.front().find(',') != std::string::npos,
                   L"listing issues exactly one listing-shaped bulk FETCH and no repair fetch",
                   *passed,
                   *failed);
        const auto countVerb = [](const ImapEndpointSnapshot& observed, std::string_view verb) noexcept
        {
            return static_cast<size_t>(std::ranges::count_if(observed.commands, [verb](const CommandRecord& command) noexcept { return command.verb == verb; }));
        };
        DebugCheck(snapshot.logins >= 1u && countVerb(snapshot, "UID SEARCH") == 1u,
                   L"listing authenticates and enumerates UIDs through libcurl",
                   *passed,
                   *failed);
        Debug::Perf::Emit(L"FileOps.Curl.Imap.TransportWitness", L"listing", listingUs, entries.size(), snapshot.fetchSets.size(), hr);

        // 2. Targeted lookup: a single UID is sent as `uid:uid` so libcurl keeps it on the
        //    listing channel instead of truncating the row after the literal.
        FilesInformationCurl::Entry selected{};
        const auto lookupStarted = std::chrono::steady_clock::now();
        hr                       = GetEntryInfo(conn, L"/INBOX/message [777-2].eml", selected);
        const uint64_t lookupUs  = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - lookupStarted).count());
        snapshot                 = server.Snapshot();
        DebugCheck(hr == S_OK && selected.sizeKnown && selected.sizeBytes == 200u && selected.name == BuildImapMessageLeafName(literalSubject, {}, 777u, 2u),
                   L"targeted lookup of a literal-subject message succeeds through libcurl",
                   *passed,
                   *failed);
        DebugCheck(! snapshot.fetchSets.empty() && snapshot.fetchSets.back() == "2:2",
                   L"targeted lookup sends the one-element range form",
                   *passed,
                   *failed);
        Debug::Perf::Emit(L"FileOps.Curl.Imap.TransportWitness", L"lookup", lookupUs, 1u, snapshot.fetchSets.size(), hr);

        // 3. libcurl contract controls (see ImapResponseCapture).
        std::string bodyCapture;
        const HRESULT bodyHr = CurlPerformImapCustomRequest(conn, L"/INBOX", "UID FETCH 1,3 (UID RFC822.SIZE)", bodyCapture, nullptr, ImapResponseCapture::Body);
        DebugCheck(bodyHr == S_OK && bodyCapture.empty(),
                   L"libcurl control: a comma-separated custom UID FETCH writes no body bytes",
                   *passed,
                   *failed);
        std::string wireCapture;
        const HRESULT wireHr =
            CurlPerformImapCustomRequest(conn, L"/INBOX", "UID FETCH 1,3 (UID RFC822.SIZE)", wireCapture, nullptr, ImapResponseCapture::WireLines);
        DebugCheck(wireHr == S_OK && wireCapture.find("* 1 FETCH (UID 1 ") != std::string::npos && wireCapture.find("* 3 FETCH (UID 3 ") != std::string::npos,
                   L"libcurl control: the same FETCH rows arrive on the wire-line channel",
                   *passed,
                   *failed);
        std::string bareBodyCapture;
        const HRESULT bareHr = CurlPerformImapCustomRequest(conn, L"/INBOX", "UID FETCH 2 (UID ENVELOPE)", bareBodyCapture, nullptr, ImapResponseCapture::Body);
        DebugCheck(bareHr == S_OK && bareBodyCapture.find(literalSubjectUtf8) != std::string::npos &&
                       bareBodyCapture.find("NIL NIL NIL NIL NIL NIL NIL NIL))") == std::string::npos,
                   L"libcurl control: a bare UID FETCH body stops after the first literal",
                   *passed,
                   *failed);

        // 4. Damaged framing is a bad response, never proof of absence.
        server.SetDamagedFetchUid(3u);
        FilesInformationCurl::Entry damaged{};
        const HRESULT damagedHr = GetEntryInfo(conn, L"/INBOX/message [777-3].eml", damaged);
        DebugCheck(damagedHr == HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP),
                   L"a truncated literal in the FETCH row reports a bad server response",
                   *passed,
                   *failed);
        FilesInformationCurl::Entry absent{};
        const HRESULT absentHr = GetEntryInfo(conn, L"/INBOX/message [777-9].eml", absent);
        DebugCheck(absentHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                   L"a well-formed reply without the UID reports the message as missing",
                   *passed,
                   *failed);

        snapshot = server.Snapshot();
        DebugCheck(SUCCEEDED(snapshot.serverHr) && snapshot.connections >= 1u, L"fake IMAP server observed no transport fault", *passed, *failed);
        server.Stop();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory exported noexcept boundary: fail the test without escaping into the host.
        Debug::Error(L"Curl IMAP transport fixture failed after std::exception.");
        DebugCheck(false, L"IMAP transport fixture must not throw", *passed, *failed);
    }
    return *failed == 0u ? S_OK : E_FAIL;
}
} // namespace FileSystemCurlInternal

#else

static_assert(true);

#endif
