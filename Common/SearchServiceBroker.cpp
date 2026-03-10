#include "SearchServiceBroker.h"

#include "Helpers.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <filesystem>
#include <limits>
#include <span>
#include <string_view>
#include <vector>

#include <sddl.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace SearchServiceBroker
{
namespace
{
constexpr uint32_t kMessageMagic        = 0x53535252u; // "RRSS"
constexpr uint32_t kMaxFrameBytes       = 16u * 1024u * 1024u;
constexpr DWORD kClientConnectTimeoutMs = 150u;

enum class MessageType : uint32_t
{
    StatusRequest  = 1u,
    StatusResponse = 2u,
    QueryRequest   = 3u,
    QueryProgress  = 4u,
    QueryBatch     = 5u,
    QueryComplete  = 6u,
    RebuildRequest = 7u,
    Ack            = 8u,
    Error          = 9u,
};

struct FrameHeader final
{
    uint32_t magic           = kMessageMagic;
    uint32_t protocolVersion = kProtocolVersion;
    uint32_t messageType     = 0u;
    uint32_t payloadBytes    = 0u;
};

struct StatusResponsePayload final
{
    uint32_t processId        = 0u;
    uint32_t pipeNameBytes    = 0u;
    uint32_t storageRootBytes = 0u;
    uint32_t reserved         = 0u;
};

struct QueryRequestPayload final
{
    uint32_t nameMode         = 0u;
    uint32_t flags            = 0u;
    uint32_t rootPathBytes    = 0u;
    uint32_t namePatternBytes = 0u;
    uint64_t maxResults       = 0u;
};

struct ProgressPayload final
{
    uint32_t phase              = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    uint32_t warningFlags       = FILESYSTEM_SEARCH_WARNING_NONE;
    int32_t statusHint          = S_OK;
    uint32_t currentPathBytes   = 0u;
    uint64_t scannedDirectories = 0u;
    uint64_t scannedFiles       = 0u;
    uint64_t candidateFiles     = 0u;
    uint64_t matchedEntries     = 0u;
};

struct CandidateBatchHeader final
{
    uint32_t count    = 0u;
    uint32_t reserved = 0u;
};

struct CandidateEntryHeader final
{
    uint32_t fileAttributes   = 0u;
    uint32_t fullPathBytes    = 0u;
    uint32_t displayNameBytes = 0u;
    uint32_t reserved         = 0u;
};

struct QueryCompletePayload final
{
    int32_t result           = S_OK;
    uint32_t fileSystemKind  = 0u;
    uint32_t flags           = 0u;
    uint64_t entryCount      = 0u;
    uint64_t fileCount       = 0u;
    uint64_t directoryCount  = 0u;
    uint64_t candidateCount  = 0u;
    uint64_t nextUsn         = 0u;
    uint64_t journalId       = 0u;
    uint64_t snapshotFileBytes      = 0u;
    uint64_t estimatedMemoryBytes   = 0u;
    uint32_t ensureReadyDurationMs  = 0u;
    uint32_t executeQueryDurationMs = 0u;
};

struct RebuildRequestPayload final
{
    uint32_t rootPathBytes = 0u;
};

struct AckPayload final
{
    int32_t result = S_OK;
};

struct ErrorPayload final
{
    int32_t result        = E_FAIL;
    uint32_t messageBytes = 0u;
};

enum QueryCompleteFlags : uint32_t
{
    QUERY_COMPLETE_FLAG_SNAPSHOT_LOADED              = 0x1u,
    QUERY_COMPLETE_FLAG_SNAPSHOT_SAVED               = 0x2u,
    QUERY_COMPLETE_FLAG_JOURNAL_AVAILABLE            = 0x4u,
    QUERY_COMPLETE_FLAG_JOURNAL_REPLAY_APPLIED       = 0x8u,
    QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_ID_MISMATCH  = 0x10u,
    QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_RANGE_INVALID = 0x20u,
    QUERY_COMPLETE_FLAG_REBUILT_SNAPSHOT_CORRUPTION  = 0x40u,
    QUERY_COMPLETE_FLAG_USED_NTFS_ENUMERATION        = 0x80u,
    QUERY_COMPLETE_FLAG_USED_TRAVERSAL_SEED          = 0x100u,
};

struct SessionContext final
{
    SessionContext() = default;
    SessionContext(const SessionContext&) = delete;
    SessionContext& operator=(const SessionContext&) = delete;
    SessionContext(SessionContext&&) = delete;
    SessionContext& operator=(SessionContext&&) = delete;

    wil::unique_handle pipe;
    LocalSearchIndexCore::Repository* repository = nullptr;
    ServerOptions options;
};

[[nodiscard]] std::wstring Utf16FromBytes(const std::byte* bytes, uint32_t byteCount) noexcept
{
    if (bytes == nullptr || byteCount == 0u)
    {
        return {};
    }

    if ((byteCount % sizeof(wchar_t)) != 0u)
    {
        return {};
    }

    const wchar_t* text = reinterpret_cast<const wchar_t*>(bytes);
    return std::wstring(text, text + (byteCount / sizeof(wchar_t)));
}

void AppendBytes(std::vector<std::byte>& buffer, const void* data, size_t byteCount)
{
    const std::byte* begin = static_cast<const std::byte*>(data);
    buffer.insert(buffer.end(), begin, begin + byteCount);
}

void AppendUtf16(std::vector<std::byte>& buffer, std::wstring_view text)
{
    if (! text.empty())
    {
        AppendBytes(buffer, text.data(), text.size() * sizeof(wchar_t));
    }
}

template<typename T>
[[nodiscard]] bool ReadPod(std::span<const std::byte>& remaining, T& out) noexcept
{
    if (remaining.size_bytes() < sizeof(T))
    {
        return false;
    }

    std::memcpy(&out, remaining.data(), sizeof(T));
    remaining = remaining.subspan(sizeof(T));
    return true;
}

[[nodiscard]] bool ReadUtf16(std::span<const std::byte>& remaining, uint32_t byteCount, std::wstring& out) noexcept
{
    if (remaining.size_bytes() < byteCount || (byteCount % sizeof(wchar_t)) != 0u)
    {
        out.clear();
        return false;
    }

    out = Utf16FromBytes(remaining.data(), byteCount);
    remaining = remaining.subspan(byteCount);
    return true;
}

HRESULT ReadExact(HANDLE handle, void* buffer, uint32_t byteCount) noexcept
{
    uint8_t* destination = static_cast<uint8_t*>(buffer);
    uint32_t totalRead   = 0u;
    while (totalRead < byteCount)
    {
        DWORD chunkRead = 0u;
        if (::ReadFile(handle, destination + totalRead, byteCount - totalRead, &chunkRead, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        if (chunkRead == 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
        }

        totalRead += chunkRead;
    }

    return S_OK;
}

HRESULT WriteExact(HANDLE handle, const void* buffer, uint32_t byteCount) noexcept
{
    const uint8_t* source = static_cast<const uint8_t*>(buffer);
    uint32_t totalWritten = 0u;
    while (totalWritten < byteCount)
    {
        DWORD chunkWritten = 0u;
        if (::WriteFile(handle, source + totalWritten, byteCount - totalWritten, &chunkWritten, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        totalWritten += chunkWritten;
    }

    return S_OK;
}

HRESULT SendFrame(HANDLE handle, MessageType messageType, uint32_t protocolVersion, const std::vector<std::byte>& payload) noexcept
{
    if (payload.size() > kMaxFrameBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }

    FrameHeader header{};
    header.protocolVersion = protocolVersion;
    header.messageType     = static_cast<uint32_t>(messageType);
    header.payloadBytes    = static_cast<uint32_t>(payload.size());

    HRESULT hr = WriteExact(handle, &header, static_cast<uint32_t>(sizeof(header)));
    if (FAILED(hr))
    {
        return hr;
    }

    if (! payload.empty())
    {
        hr = WriteExact(handle, payload.data(), static_cast<uint32_t>(payload.size()));
    }
    return hr;
}

HRESULT ReceiveFrame(HANDLE handle, FrameHeader& outHeader, std::vector<std::byte>& outPayload) noexcept
{
    outHeader = {};
    outPayload.clear();

    HRESULT hr = ReadExact(handle, &outHeader, static_cast<uint32_t>(sizeof(outHeader)));
    if (FAILED(hr))
    {
        return hr;
    }

    if (outHeader.magic != kMessageMagic)
    {
        return RPC_S_PROTOCOL_ERROR;
    }

    if (outHeader.payloadBytes > kMaxFrameBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }

    outPayload.resize(outHeader.payloadBytes);
    if (outHeader.payloadBytes != 0u)
    {
        hr = ReadExact(handle, outPayload.data(), outHeader.payloadBytes);
    }
    return hr;
}

[[nodiscard]] std::wstring BuildPipeName(std::wstring_view configuredName) noexcept
{
    if (! configuredName.empty())
    {
        return std::wstring(configuredName);
    }

    return std::wstring(LR"(\\.\pipe\RedSalamander.SearchService.v1)");
}

[[nodiscard]] std::wstring GetEnvironmentValue(std::wstring_view name) noexcept
{
    std::wstring key(name);
    const DWORD required = ::GetEnvironmentVariableW(key.c_str(), nullptr, 0u);
    if (required == 0u)
    {
        return {};
    }

    std::wstring value(static_cast<size_t>(required), L'\0');
    const DWORD written = ::GetEnvironmentVariableW(key.c_str(), value.data(), required);
    if (written == 0u)
    {
        return {};
    }

    if (! value.empty() && value.back() == L'\0')
    {
        value.pop_back();
    }
    return value;
}

HRESULT ConnectClientPipe(wil::unique_handle& outPipe) noexcept
{
    const std::wstring pipeName = GetConfiguredPipeName();
    const ULONGLONG deadline = ::GetTickCount64() + 750u;

    for (;;)
    {
        if (::WaitNamedPipeW(pipeName.c_str(), kClientConnectTimeoutMs) == 0)
        {
            const DWORD waitError = ::GetLastError();
            if (waitError != ERROR_SEM_TIMEOUT && waitError != ERROR_FILE_NOT_FOUND && waitError != ERROR_PIPE_BUSY)
            {
                return HRESULT_FROM_WIN32(waitError);
            }
        }

        outPipe.reset(::CreateFileW(pipeName.c_str(),
                                    GENERIC_READ | GENERIC_WRITE,
                                    0u,
                                    nullptr,
                                    OPEN_EXISTING,
                                    FILE_ATTRIBUTE_NORMAL,
                                    nullptr));
        if (outPipe)
        {
            return S_OK;
        }

        const DWORD error = ::GetLastError();
        if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PIPE_BUSY && error != ERROR_SEM_TIMEOUT)
        {
            return HRESULT_FROM_WIN32(error);
        }

        if (::GetTickCount64() >= deadline)
        {
            return HRESULT_FROM_WIN32(error);
        }

        ::Sleep(25u);
    }
}

HRESULT SendProtocolError(HANDLE handle, uint32_t protocolVersion, HRESULT hr, std::wstring_view message) noexcept
{
    ErrorPayload payload{};
    payload.result       = hr;
    payload.messageBytes = static_cast<uint32_t>(message.size() * sizeof(wchar_t));

    std::vector<std::byte> buffer;
    buffer.reserve(sizeof(payload) + payload.messageBytes);
    AppendBytes(buffer, &payload, sizeof(payload));
    AppendUtf16(buffer, message);
    return SendFrame(handle, MessageType::Error, protocolVersion, buffer);
}

HRESULT SendStatusResponse(HANDLE handle, uint32_t protocolVersion, const ServerOptions& options) noexcept
{
    const std::wstring pipeName    = BuildPipeName(options.pipeName);
    const std::wstring storageRoot = options.storageRootDirectory;

    StatusResponsePayload payload{};
    payload.processId        = ::GetCurrentProcessId();
    payload.pipeNameBytes    = static_cast<uint32_t>(pipeName.size() * sizeof(wchar_t));
    payload.storageRootBytes = static_cast<uint32_t>(storageRoot.size() * sizeof(wchar_t));

    std::vector<std::byte> buffer;
    buffer.reserve(sizeof(payload) + payload.pipeNameBytes + payload.storageRootBytes);
    AppendBytes(buffer, &payload, sizeof(payload));
    AppendUtf16(buffer, pipeName);
    AppendUtf16(buffer, storageRoot);
    return SendFrame(handle, MessageType::StatusResponse, protocolVersion, buffer);
}

HRESULT SendProgress(HANDLE handle, uint32_t protocolVersion, const QueryProgress& progress) noexcept
{
    ProgressPayload payload{};
    payload.phase              = static_cast<uint32_t>(progress.phase);
    payload.warningFlags       = progress.warningFlags;
    payload.statusHint         = progress.statusHint;
    payload.currentPathBytes   = static_cast<uint32_t>(progress.currentPath.size() * sizeof(wchar_t));
    payload.scannedDirectories = progress.scannedDirectories;
    payload.scannedFiles       = progress.scannedFiles;
    payload.candidateFiles     = progress.candidateFiles;
    payload.matchedEntries     = progress.matchedEntries;

    std::vector<std::byte> buffer;
    buffer.reserve(sizeof(payload) + payload.currentPathBytes);
    AppendBytes(buffer, &payload, sizeof(payload));
    AppendUtf16(buffer, progress.currentPath);
    return SendFrame(handle, MessageType::QueryProgress, protocolVersion, buffer);
}

HRESULT SendCandidates(HANDLE handle,
                       uint32_t protocolVersion,
                       std::span<const LocalSearchIndexCore::Candidate> candidates,
                       uint32_t disconnectAfterBatches,
                       uint32_t& batchesSent) noexcept
{
    constexpr size_t kBatchItems = 256u;

    for (size_t offset = 0; offset < candidates.size(); offset += kBatchItems)
    {
        const size_t count = std::min(kBatchItems, candidates.size() - offset);

        CandidateBatchHeader batchHeader{};
        batchHeader.count = static_cast<uint32_t>(count);

        std::vector<std::byte> buffer;
        buffer.reserve(sizeof(batchHeader) + (count * sizeof(CandidateEntryHeader)));
        AppendBytes(buffer, &batchHeader, sizeof(batchHeader));

        for (size_t index = 0; index < count; ++index)
        {
            const auto& candidate = candidates[offset + index];
            CandidateEntryHeader entry{};
            entry.fileAttributes   = candidate.fileAttributes;
            entry.fullPathBytes    = static_cast<uint32_t>(candidate.fullPath.size() * sizeof(wchar_t));
            entry.displayNameBytes = static_cast<uint32_t>(candidate.displayName.size() * sizeof(wchar_t));
            AppendBytes(buffer, &entry, sizeof(entry));
            AppendUtf16(buffer, candidate.fullPath);
            AppendUtf16(buffer, candidate.displayName);
        }

        HRESULT hr = SendFrame(handle, MessageType::QueryBatch, protocolVersion, buffer);
        if (FAILED(hr))
        {
            return hr;
        }

        ++batchesSent;
        if (disconnectAfterBatches != 0u && batchesSent >= disconnectAfterBatches)
        {
            return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
        }
    }

    return S_OK;
}

uint32_t PackQueryStatsFlags(const LocalSearchIndexCore::QueryStats& stats) noexcept
{
    uint32_t flags = 0u;
    if (stats.snapshotLoaded) { flags |= QUERY_COMPLETE_FLAG_SNAPSHOT_LOADED; }
    if (stats.snapshotSaved) { flags |= QUERY_COMPLETE_FLAG_SNAPSHOT_SAVED; }
    if (stats.journalAvailable) { flags |= QUERY_COMPLETE_FLAG_JOURNAL_AVAILABLE; }
    if (stats.journalReplayApplied) { flags |= QUERY_COMPLETE_FLAG_JOURNAL_REPLAY_APPLIED; }
    if (stats.rebuiltJournalIdMismatch) { flags |= QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_ID_MISMATCH; }
    if (stats.rebuiltJournalRangeInvalid) { flags |= QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_RANGE_INVALID; }
    if (stats.rebuiltSnapshotCorruption) { flags |= QUERY_COMPLETE_FLAG_REBUILT_SNAPSHOT_CORRUPTION; }
    if (stats.usedNtfsEnumeration) { flags |= QUERY_COMPLETE_FLAG_USED_NTFS_ENUMERATION; }
    if (stats.usedTraversalSeed) { flags |= QUERY_COMPLETE_FLAG_USED_TRAVERSAL_SEED; }
    return flags;
}

void UnpackQueryStatsFlags(uint32_t flags, LocalSearchIndexCore::QueryStats& stats) noexcept
{
    stats.snapshotLoaded             = (flags & QUERY_COMPLETE_FLAG_SNAPSHOT_LOADED) != 0u;
    stats.snapshotSaved              = (flags & QUERY_COMPLETE_FLAG_SNAPSHOT_SAVED) != 0u;
    stats.journalAvailable           = (flags & QUERY_COMPLETE_FLAG_JOURNAL_AVAILABLE) != 0u;
    stats.journalReplayApplied       = (flags & QUERY_COMPLETE_FLAG_JOURNAL_REPLAY_APPLIED) != 0u;
    stats.rebuiltJournalIdMismatch   = (flags & QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_ID_MISMATCH) != 0u;
    stats.rebuiltJournalRangeInvalid = (flags & QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_RANGE_INVALID) != 0u;
    stats.rebuiltSnapshotCorruption  = (flags & QUERY_COMPLETE_FLAG_REBUILT_SNAPSHOT_CORRUPTION) != 0u;
    stats.usedNtfsEnumeration        = (flags & QUERY_COMPLETE_FLAG_USED_NTFS_ENUMERATION) != 0u;
    stats.usedTraversalSeed          = (flags & QUERY_COMPLETE_FLAG_USED_TRAVERSAL_SEED) != 0u;
}

HRESULT SendQueryComplete(HANDLE handle,
                          uint32_t protocolVersion,
                          HRESULT result,
                          const LocalSearchIndexCore::QueryStats& stats) noexcept
{
    QueryCompletePayload payload{};
    payload.result          = result;
    payload.fileSystemKind  = static_cast<uint32_t>(stats.fileSystemKind);
    payload.flags           = PackQueryStatsFlags(stats);
    payload.entryCount      = stats.entryCount;
    payload.fileCount       = stats.fileCount;
    payload.directoryCount  = stats.directoryCount;
    payload.candidateCount  = stats.candidateCount;
    payload.nextUsn         = stats.nextUsn;
    payload.journalId       = stats.journalId;
    payload.snapshotFileBytes      = stats.snapshotFileBytes;
    payload.estimatedMemoryBytes   = stats.estimatedMemoryBytes;
    payload.ensureReadyDurationMs  = stats.ensureReadyDurationMs;
    payload.executeQueryDurationMs = stats.executeQueryDurationMs;

    std::vector<std::byte> buffer;
    buffer.reserve(sizeof(payload));
    AppendBytes(buffer, &payload, sizeof(payload));
    return SendFrame(handle, MessageType::QueryComplete, protocolVersion, buffer);
}

HRESULT SendAck(HANDLE handle, uint32_t protocolVersion, HRESULT hr) noexcept
{
    AckPayload payload{};
    payload.result = hr;

    std::vector<std::byte> buffer;
    buffer.reserve(sizeof(payload));
    AppendBytes(buffer, &payload, sizeof(payload));
    return SendFrame(handle, MessageType::Ack, protocolVersion, buffer);
}

HRESULT STDMETHODCALLTYPE ServerCancelCheck(void* cookie) noexcept
{
    UNREFERENCED_PARAMETER(cookie);
    return S_OK;
}

HRESULT HandleStatusRequest(SessionContext& session) noexcept
{
    return SendStatusResponse(session.pipe.get(), session.options.protocolVersion, session.options);
}

HRESULT HandleRebuildRequest(SessionContext& session, std::span<const std::byte> payloadBytes) noexcept
{
    if (! session.options.allowRebuildRequests)
    {
        return SendAck(session.pipe.get(), session.options.protocolVersion, E_ACCESSDENIED);
    }

    RebuildRequestPayload payload{};
    if (! ReadPod(payloadBytes, payload))
    {
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, RPC_S_PROTOCOL_ERROR, L"Invalid rebuild request.");
    }

    std::wstring rootPath;
    if (! ReadUtf16(payloadBytes, payload.rootPathBytes, rootPath))
    {
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, RPC_S_PROTOCOL_ERROR, L"Invalid rebuild path.");
    }

    const HRESULT hr = session.repository->InvalidateRoot(rootPath, true);
    Debug::Info(L"SearchServiceBroker: rebuild request root='{}' hr=0x{:08X}", rootPath, static_cast<unsigned long>(hr));
    return SendAck(session.pipe.get(), session.options.protocolVersion, hr);
}

HRESULT HandleQueryRequest(SessionContext& session, std::span<const std::byte> payloadBytes) noexcept
{
    QueryRequestPayload payload{};
    if (! ReadPod(payloadBytes, payload))
    {
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, RPC_S_PROTOCOL_ERROR, L"Invalid query request.");
    }

    std::wstring rootPath;
    std::wstring namePattern;
    if (! ReadUtf16(payloadBytes, payload.rootPathBytes, rootPath) || ! ReadUtf16(payloadBytes, payload.namePatternBytes, namePattern))
    {
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, RPC_S_PROTOCOL_ERROR, L"Invalid query strings.");
    }

    LocalSearchIndexCore::QueryPlan plan{};
    plan.rootPath           = std::move(rootPath);
    plan.namePattern        = std::move(namePattern);
    plan.nameMode           = static_cast<FileSystemSearchNameMode>(payload.nameMode);
    plan.matchCaseName      = (payload.flags & FILESYSTEM_SEARCH_MATCH_CASE_NAME) != 0;
    plan.recursive          = (payload.flags & FILESYSTEM_SEARCH_RECURSIVE) != 0;
    plan.includeFiles       = (payload.flags & FILESYSTEM_SEARCH_INCLUDE_FILES) != 0;
    plan.includeDirectories = (payload.flags & FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES) != 0;
    plan.maxResults         = payload.maxResults;

    QueryProgress progress{};
    progress.phase       = FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP;
    progress.currentPath = plan.rootPath;

    HRESULT hr = SendProgress(session.pipe.get(), session.options.protocolVersion, progress);
    if (FAILED(hr))
    {
        return hr;
    }

    LocalSearchIndexCore::QueryStats stats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    hr = session.repository->Query(plan, &ServerCancelCheck, nullptr, candidates, &stats);

    progress.phase              = FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP;
    progress.statusHint         = hr;
    progress.scannedDirectories = stats.directoryCount;
    progress.scannedFiles       = stats.fileCount;
    progress.candidateFiles     = stats.candidateCount;
    const HRESULT progressHr = SendProgress(session.pipe.get(), session.options.protocolVersion, progress);
    if (FAILED(progressHr))
    {
        return progressHr;
    }

    if (FAILED(hr))
    {
        Debug::Warning(L"SearchServiceBroker: query failed root='{}' pattern='{}' mode={} hr=0x{:08X} candidates={} snapshotBytes={} memoryBytes={} "
                       L"readyMs={} queryMs={}",
                       plan.rootPath,
                       plan.namePattern,
                       static_cast<uint32_t>(plan.nameMode),
                       static_cast<unsigned long>(hr),
                       stats.candidateCount,
                       stats.snapshotFileBytes,
                       stats.estimatedMemoryBytes,
                       stats.ensureReadyDurationMs,
                       stats.executeQueryDurationMs);
        return SendQueryComplete(session.pipe.get(), session.options.protocolVersion, hr, stats);
    }

    uint32_t batchesSent = 0u;
    hr = SendCandidates(session.pipe.get(), session.options.protocolVersion, candidates, session.options.disconnectAfterBatches, batchesSent);
    if (FAILED(hr))
    {
        return hr;
    }

    Debug::Info(L"SearchServiceBroker: query root='{}' pattern='{}' mode={} recursive={} includeFiles={} includeDirs={} candidates={} "
                L"snapshotBytes={} memoryBytes={} readyMs={} queryMs={} batches={}",
                plan.rootPath,
                plan.namePattern,
                static_cast<uint32_t>(plan.nameMode),
                plan.recursive,
                plan.includeFiles,
                plan.includeDirectories,
                stats.candidateCount,
                stats.snapshotFileBytes,
                stats.estimatedMemoryBytes,
                stats.ensureReadyDurationMs,
                stats.executeQueryDurationMs,
                batchesSent);
    return SendQueryComplete(session.pipe.get(), session.options.protocolVersion, S_OK, stats);
}

HRESULT HandleClient(SessionContext& session) noexcept
{
    FrameHeader header{};
    std::vector<std::byte> payload;
    HRESULT hr = ReceiveFrame(session.pipe.get(), header, payload);
    if (FAILED(hr))
    {
        return hr;
    }

    if (header.protocolVersion != session.options.protocolVersion)
    {
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, RPC_S_PROTOCOL_ERROR, L"Protocol version mismatch.");
    }

    std::span<const std::byte> payloadBytes(payload.data(), payload.size());
    switch (static_cast<MessageType>(header.messageType))
    {
        case MessageType::StatusRequest: return HandleStatusRequest(session);
        case MessageType::QueryRequest: return HandleQueryRequest(session, payloadBytes);
        case MessageType::RebuildRequest: return HandleRebuildRequest(session, payloadBytes);
        case MessageType::StatusResponse:
        case MessageType::QueryProgress:
        case MessageType::QueryBatch:
        case MessageType::QueryComplete:
        case MessageType::Ack:
        case MessageType::Error:
        default: return SendProtocolError(session.pipe.get(), session.options.protocolVersion, RPC_S_PROTOCOL_ERROR, L"Unsupported request.");
    }
}

HRESULT CreatePipeSecurity(SECURITY_ATTRIBUTES& outAttributes, wil::unique_hlocal& outDescriptor) noexcept
{
    outAttributes = {};
    outDescriptor.reset();

    static constexpr wchar_t kPipeSddl[] = L"D:P(A;;GA;;;SY)(A;;GA;;;BA)(A;;GRGW;;;IU)";

    PSECURITY_DESCRIPTOR descriptor = nullptr;
    if (::ConvertStringSecurityDescriptorToSecurityDescriptorW(kPipeSddl, SDDL_REVISION_1, &descriptor, nullptr) == 0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    outDescriptor.reset(descriptor);
    outAttributes.nLength              = sizeof(outAttributes);
    outAttributes.lpSecurityDescriptor = outDescriptor.get();
    outAttributes.bInheritHandle       = FALSE;
    return S_OK;
}

HRESULT CreateServerPipe(const std::wstring& pipeName, SECURITY_ATTRIBUTES& securityAttributes, wil::unique_handle& outPipe) noexcept
{
    outPipe.reset(::CreateNamedPipeW(pipeName.c_str(),
                                     PIPE_ACCESS_DUPLEX | FILE_FLAG_OVERLAPPED,
                                     PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT | PIPE_REJECT_REMOTE_CLIENTS,
                                     PIPE_UNLIMITED_INSTANCES,
                                     256u * 1024u,
                                     256u * 1024u,
                                     0u,
                                     &securityAttributes));
    if (! outPipe)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    return S_OK;
}

HRESULT WaitForClientConnection(HANDLE pipe, HANDLE stopEvent) noexcept
{
    wil::unique_event connectEvent;
    connectEvent.create();
    if (! connectEvent)
    {
        return E_OUTOFMEMORY;
    }

    OVERLAPPED overlapped{};
    overlapped.hEvent = connectEvent.get();

    if (::ConnectNamedPipe(pipe, &overlapped) != 0)
    {
        return S_OK;
    }

    const DWORD error = ::GetLastError();
    if (error == ERROR_PIPE_CONNECTED)
    {
        return S_OK;
    }

    if (error != ERROR_IO_PENDING)
    {
        return HRESULT_FROM_WIN32(error);
    }

    HANDLE waits[2] = {stopEvent, connectEvent.get()};
    const DWORD waitCount = stopEvent ? 2u : 1u;
    const DWORD waitResult = ::WaitForMultipleObjects(waitCount, stopEvent ? waits : waits + 1, FALSE, INFINITE);
    if (stopEvent && waitResult == WAIT_OBJECT_0)
    {
        static_cast<void>(::CancelIoEx(pipe, &overlapped));
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    DWORD bytesTransferred = 0u;
    if (::GetOverlappedResult(pipe, &overlapped, &bytesTransferred, FALSE) == 0)
    {
        const DWORD resultError = ::GetLastError();
        if (resultError == ERROR_PIPE_CONNECTED)
        {
            return S_OK;
        }

        return HRESULT_FROM_WIN32(resultError);
    }

    return S_OK;
}
} // namespace

std::wstring GetDefaultPipeName() noexcept
{
    return BuildPipeName({});
}

std::wstring GetConfiguredPipeName() noexcept
{
    const std::wstring envPipe = GetEnvironmentValue(kPipeNameEnvVar);
    return BuildPipeName(envPipe);
}

std::wstring GetProgramDataSearchIndexRoot() noexcept
{
    const std::wstring root = GetEnvironmentValue(L"ProgramData");
    if (root.empty())
    {
        return {};
    }

    return (std::filesystem::path(root) / L"RedSalamander" / L"SearchIndex").wstring();
}

HRESULT GetStatus(ServiceStatus& outStatus) noexcept
{
    try
    {
        outStatus = {};

        wil::unique_handle pipe;
        HRESULT hr = ConnectClientPipe(pipe);
        if (FAILED(hr))
        {
            return hr;
        }

        std::vector<std::byte> emptyPayload;
        hr = SendFrame(pipe.get(), MessageType::StatusRequest, kProtocolVersion, emptyPayload);
        if (FAILED(hr))
        {
            return hr;
        }

        FrameHeader header{};
        std::vector<std::byte> payload;
        hr = ReceiveFrame(pipe.get(), header, payload);
        if (FAILED(hr))
        {
            return hr;
        }

        if (header.protocolVersion != kProtocolVersion)
        {
            return RPC_S_PROTOCOL_ERROR;
        }

        std::span<const std::byte> remaining(payload.data(), payload.size());
        if (static_cast<MessageType>(header.messageType) == MessageType::Error)
        {
            ErrorPayload error{};
            return ReadPod(remaining, error) ? static_cast<HRESULT>(error.result) : RPC_S_PROTOCOL_ERROR;
        }

        if (static_cast<MessageType>(header.messageType) != MessageType::StatusResponse)
        {
            return RPC_S_PROTOCOL_ERROR;
        }

        StatusResponsePayload statusPayload{};
        if (! ReadPod(remaining, statusPayload) ||
            ! ReadUtf16(remaining, statusPayload.pipeNameBytes, outStatus.pipeName) ||
            ! ReadUtf16(remaining, statusPayload.storageRootBytes, outStatus.storageRootDirectory))
        {
            outStatus = {};
            return RPC_S_PROTOCOL_ERROR;
        }

        outStatus.protocolVersion = header.protocolVersion;
        outStatus.processId       = statusPayload.processId;
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SearchServiceBroker: GetStatus failed with an unexpected std::exception.");
        outStatus = {};
        return E_FAIL;
    }
}

HRESULT Query(const QueryRequest& request,
              ProgressCallbackFn progressCallback,
              void* progressCookie,
              LocalSearchIndexCore::CancelCheckFn cancelCheck,
              void* cancelCookie,
              std::vector<LocalSearchIndexCore::Candidate>& outCandidates,
              LocalSearchIndexCore::QueryStats* outStats) noexcept
{
    try
    {
        outCandidates.clear();
        if (outStats != nullptr)
        {
            *outStats = {};
        }

        if (request.rootPath.empty())
        {
            return E_INVALIDARG;
        }

        wil::unique_handle pipe;
        HRESULT hr = ConnectClientPipe(pipe);
        if (FAILED(hr))
        {
            return hr;
        }

        QueryRequestPayload payload{};
        payload.nameMode         = static_cast<uint32_t>(request.nameMode);
        payload.flags            = static_cast<uint32_t>(request.flags);
        payload.rootPathBytes    = static_cast<uint32_t>(request.rootPath.size() * sizeof(wchar_t));
        payload.namePatternBytes = static_cast<uint32_t>(request.namePattern.size() * sizeof(wchar_t));
        payload.maxResults       = request.maxResults;

        std::vector<std::byte> requestBuffer;
        requestBuffer.reserve(sizeof(payload) + payload.rootPathBytes + payload.namePatternBytes);
        AppendBytes(requestBuffer, &payload, sizeof(payload));
        AppendUtf16(requestBuffer, request.rootPath);
        AppendUtf16(requestBuffer, request.namePattern);

        hr = SendFrame(pipe.get(), MessageType::QueryRequest, kProtocolVersion, requestBuffer);
        if (FAILED(hr))
        {
            return hr;
        }

        for (;;)
        {
            if (cancelCheck != nullptr)
            {
                hr = cancelCheck(cancelCookie);
                if (FAILED(hr))
                {
                    return hr;
                }
            }

            FrameHeader header{};
            std::vector<std::byte> payloadBytes;
            hr = ReceiveFrame(pipe.get(), header, payloadBytes);
            if (FAILED(hr))
            {
                return hr;
            }

            if (header.protocolVersion != kProtocolVersion)
            {
                return RPC_S_PROTOCOL_ERROR;
            }

            std::span<const std::byte> remaining(payloadBytes.data(), payloadBytes.size());
            switch (static_cast<MessageType>(header.messageType))
            {
                case MessageType::QueryProgress:
                {
                    ProgressPayload progressPayload{};
                    QueryProgress progress{};
                    if (! ReadPod(remaining, progressPayload) ||
                        ! ReadUtf16(remaining, progressPayload.currentPathBytes, progress.currentPath))
                    {
                        return RPC_S_PROTOCOL_ERROR;
                    }

                    progress.phase              = static_cast<FileSystemSearchPhase>(progressPayload.phase);
                    progress.warningFlags       = progressPayload.warningFlags;
                    progress.statusHint         = progressPayload.statusHint;
                    progress.scannedDirectories = progressPayload.scannedDirectories;
                    progress.scannedFiles       = progressPayload.scannedFiles;
                    progress.candidateFiles     = progressPayload.candidateFiles;
                    progress.matchedEntries     = progressPayload.matchedEntries;

                    if (progressCallback != nullptr)
                    {
                        hr = progressCallback(&progress, progressCookie);
                        if (FAILED(hr))
                        {
                            return hr;
                        }
                    }
                    break;
                }

                case MessageType::QueryBatch:
                {
                    CandidateBatchHeader batchHeader{};
                    if (! ReadPod(remaining, batchHeader))
                    {
                        return RPC_S_PROTOCOL_ERROR;
                    }

                    for (uint32_t index = 0u; index < batchHeader.count; ++index)
                    {
                        CandidateEntryHeader entryHeader{};
                        LocalSearchIndexCore::Candidate candidate{};
                        if (! ReadPod(remaining, entryHeader) ||
                            ! ReadUtf16(remaining, entryHeader.fullPathBytes, candidate.fullPath) ||
                            ! ReadUtf16(remaining, entryHeader.displayNameBytes, candidate.displayName))
                        {
                            return RPC_S_PROTOCOL_ERROR;
                        }

                        candidate.fileAttributes = entryHeader.fileAttributes;
                        outCandidates.push_back(std::move(candidate));
                    }
                    break;
                }

                case MessageType::QueryComplete:
                {
                    QueryCompletePayload complete{};
                    if (! ReadPod(remaining, complete))
                    {
                        return RPC_S_PROTOCOL_ERROR;
                    }

                    if (outStats != nullptr)
                    {
                        outStats->fileSystemKind = static_cast<LocalSearchIndexCore::FileSystemKind>(complete.fileSystemKind);
                        outStats->entryCount     = complete.entryCount;
                        outStats->fileCount      = complete.fileCount;
                        outStats->directoryCount = complete.directoryCount;
                        outStats->candidateCount = complete.candidateCount;
                        outStats->nextUsn        = complete.nextUsn;
                        outStats->journalId      = complete.journalId;
                        outStats->snapshotFileBytes      = complete.snapshotFileBytes;
                        outStats->estimatedMemoryBytes   = complete.estimatedMemoryBytes;
                        outStats->ensureReadyDurationMs  = complete.ensureReadyDurationMs;
                        outStats->executeQueryDurationMs = complete.executeQueryDurationMs;
                        UnpackQueryStatsFlags(complete.flags, *outStats);
                    }

                    return static_cast<HRESULT>(complete.result);
                }

                case MessageType::Error:
                {
                    ErrorPayload error{};
                    return ReadPod(remaining, error) ? static_cast<HRESULT>(error.result) : RPC_S_PROTOCOL_ERROR;
                }

                case MessageType::StatusRequest:
                case MessageType::StatusResponse:
                case MessageType::QueryRequest:
                case MessageType::RebuildRequest:
                case MessageType::Ack:
                default: return RPC_S_PROTOCOL_ERROR;
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SearchServiceBroker: Query failed with an unexpected std::exception.");
        outCandidates.clear();
        if (outStats != nullptr)
        {
            *outStats = {};
        }
        return E_FAIL;
    }
}

HRESULT RequestRebuild(std::wstring_view rootPath) noexcept
{
    try
    {
        wil::unique_handle pipe;
        HRESULT hr = ConnectClientPipe(pipe);
        if (FAILED(hr))
        {
            return hr;
        }

        RebuildRequestPayload payload{};
        payload.rootPathBytes = static_cast<uint32_t>(rootPath.size() * sizeof(wchar_t));

        std::vector<std::byte> buffer;
        buffer.reserve(sizeof(payload) + payload.rootPathBytes);
        AppendBytes(buffer, &payload, sizeof(payload));
        AppendUtf16(buffer, rootPath);

        hr = SendFrame(pipe.get(), MessageType::RebuildRequest, kProtocolVersion, buffer);
        if (FAILED(hr))
        {
            return hr;
        }

        FrameHeader header{};
        std::vector<std::byte> payloadBytes;
        hr = ReceiveFrame(pipe.get(), header, payloadBytes);
        if (FAILED(hr))
        {
            return hr;
        }

        if (header.protocolVersion != kProtocolVersion)
        {
            return RPC_S_PROTOCOL_ERROR;
        }

        std::span<const std::byte> remaining(payloadBytes.data(), payloadBytes.size());
        if (static_cast<MessageType>(header.messageType) == MessageType::Ack)
        {
            AckPayload ack{};
            return ReadPod(remaining, ack) ? static_cast<HRESULT>(ack.result) : RPC_S_PROTOCOL_ERROR;
        }

        if (static_cast<MessageType>(header.messageType) == MessageType::Error)
        {
            ErrorPayload error{};
            return ReadPod(remaining, error) ? static_cast<HRESULT>(error.result) : RPC_S_PROTOCOL_ERROR;
        }

        return RPC_S_PROTOCOL_ERROR;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SearchServiceBroker: RequestRebuild failed with an unexpected std::exception.");
        return E_FAIL;
    }
}

HRESULT RunServer(const ServerOptions& options, HANDLE stopEvent, ServerRunResult* outResult) noexcept
{
    try
    {
        ServerRunResult result{};

        ServerOptions effectiveOptions = options;
        effectiveOptions.pipeName = BuildPipeName(options.pipeName);
        if (effectiveOptions.storageRootDirectory.empty())
        {
            effectiveOptions.storageRootDirectory = GetProgramDataSearchIndexRoot();
        }

        Debug::Info(L"SearchServiceBroker: server start pipe='{}' storage='{}' protocol={} allowRebuild={} maxRequests={} disconnectAfterBatches={}",
                    effectiveOptions.pipeName,
                    effectiveOptions.storageRootDirectory,
                    effectiveOptions.protocolVersion,
                    effectiveOptions.allowRebuildRequests,
                    effectiveOptions.maxRequestsBeforeExit,
                    effectiveOptions.disconnectAfterBatches);

        LocalSearchIndexCore::Repository repository({.snapshotRootDirectory = effectiveOptions.storageRootDirectory});

        SECURITY_ATTRIBUTES securityAttributes{};
        wil::unique_hlocal securityDescriptor;
        HRESULT hr = CreatePipeSecurity(securityAttributes, securityDescriptor);
        if (FAILED(hr))
        {
            return hr;
        }

        for (;;)
        {
            if (stopEvent != nullptr && ::WaitForSingleObject(stopEvent, 0u) == WAIT_OBJECT_0)
            {
                break;
            }

            wil::unique_handle pipe;
            hr = CreateServerPipe(effectiveOptions.pipeName, securityAttributes, pipe);
            if (FAILED(hr))
            {
                return hr;
            }

            hr = WaitForClientConnection(pipe.get(), stopEvent);
            if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                break;
            }
            if (FAILED(hr))
            {
                return hr;
            }

            SessionContext session{};
            session.pipe       = std::move(pipe);
            session.repository = &repository;
            session.options    = effectiveOptions;

            hr = HandleClient(session);
            if (FAILED(hr) && hr != HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE))
            {
                return hr;
            }

            ++result.handledRequests;
            static_cast<void>(::FlushFileBuffers(session.pipe.get()));
            static_cast<void>(::DisconnectNamedPipe(session.pipe.get()));

            if (effectiveOptions.maxRequestsBeforeExit != 0u && result.handledRequests >= effectiveOptions.maxRequestsBeforeExit)
            {
                break;
            }
        }

        if (outResult != nullptr)
        {
            *outResult = result;
        }

        Debug::Info(L"SearchServiceBroker: server stop pipe='{}' handledRequests={}", effectiveOptions.pipeName, result.handledRequests);
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SearchServiceBroker: RunServer failed with an unexpected std::exception.");
        if (outResult != nullptr)
        {
            *outResult = {};
        }
        return E_FAIL;
    }
}
} // namespace SearchServiceBroker
