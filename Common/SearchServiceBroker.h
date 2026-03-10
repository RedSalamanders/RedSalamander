#pragma once

#include <cstdint>
#include <string>
#include <vector>

#include "LocalSearchIndexCore.h"
#include "PlugInterfaces/FileSystem.h"

namespace SearchServiceBroker
{
inline constexpr uint32_t kProtocolVersion = 1u;
inline constexpr wchar_t kServiceName[]    = L"RedSalamanderSearchService";
inline constexpr wchar_t kPipeNameEnvVar[] = L"REDSALAMANDER_SEARCH_SERVICE_PIPE";

struct ServiceStatus final
{
    uint32_t protocolVersion = 0u;
    uint32_t processId       = 0u;
    std::wstring pipeName;
    std::wstring storageRootDirectory;
};

struct QueryRequest final
{
    std::wstring rootPath;
    std::wstring namePattern;
    FileSystemSearchNameMode nameMode = FILESYSTEM_SEARCH_NAME_DISABLED;
    FileSystemSearchFlags flags       = FILESYSTEM_SEARCH_NONE;
    bool recursive                    = false;
    bool includeFiles                 = false;
    bool includeDirectories           = false;
    bool matchCaseName                = false;
    uint64_t maxResults               = 0u;
};

struct QueryProgress final
{
    FileSystemSearchPhase phase = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    uint32_t warningFlags       = FILESYSTEM_SEARCH_WARNING_NONE;
    HRESULT statusHint          = S_OK;
    uint64_t scannedDirectories = 0u;
    uint64_t scannedFiles       = 0u;
    uint64_t candidateFiles     = 0u;
    uint64_t matchedEntries     = 0u;
    std::wstring currentPath;
};

using ProgressCallbackFn = HRESULT(STDMETHODCALLTYPE*)(const QueryProgress* progress, void* cookie) noexcept;

struct ServerOptions final
{
    std::wstring pipeName;
    std::wstring storageRootDirectory;
    uint32_t protocolVersion          = kProtocolVersion;
    uint32_t maxRequestsBeforeExit    = 0u;
    uint32_t disconnectAfterBatches   = 0u;
    bool allowRebuildRequests         = true;
};

struct ServerRunResult final
{
    uint32_t handledRequests = 0u;
};

[[nodiscard]] std::wstring GetDefaultPipeName() noexcept;
[[nodiscard]] std::wstring GetConfiguredPipeName() noexcept;
[[nodiscard]] std::wstring GetProgramDataSearchIndexRoot() noexcept;

HRESULT GetStatus(ServiceStatus& outStatus) noexcept;
HRESULT Query(const QueryRequest& request,
              ProgressCallbackFn progressCallback,
              void* progressCookie,
              LocalSearchIndexCore::CancelCheckFn cancelCheck,
              void* cancelCookie,
              std::vector<LocalSearchIndexCore::Candidate>& outCandidates,
              LocalSearchIndexCore::QueryStats* outStats) noexcept;
HRESULT RequestRebuild(std::wstring_view rootPath) noexcept;

HRESULT RunServer(const ServerOptions& options, HANDLE stopEvent, ServerRunResult* outResult) noexcept;
} // namespace SearchServiceBroker
