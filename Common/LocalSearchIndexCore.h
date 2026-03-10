#pragma once

#include <cstdint>
#include <memory>
#include <mutex>
#include <regex>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "PlugInterfaces/FileSystem.h"

namespace LocalSearchIndexCore
{
struct VolumeIndex;

struct RepositoryOptions final
{
    std::wstring snapshotRootDirectory;
};

enum class FileSystemKind : uint32_t
{
    Unsupported = 0,
    Ntfs        = 1,
    Refs        = 2,
};

struct SupportInfo final
{
    FileSystemKind fileSystemKind = FileSystemKind::Unsupported;
    bool indexable                = false;
    std::wstring normalizedRootPath;
};

struct QueryPlan final
{
    std::wstring rootPath;
    std::wstring namePattern;
    FileSystemSearchNameMode nameMode   = FILESYSTEM_SEARCH_NAME_DISABLED;
    const std::wregex* compiledNameRegex = nullptr;
    bool matchCaseName                  = false;
    bool recursive                      = false;
    bool includeFiles                   = false;
    bool includeDirectories             = false;
    uint64_t maxResults                 = 0u;
};

struct Candidate final
{
    std::wstring fullPath;
    std::wstring displayName;
    unsigned long fileAttributes = 0u;
};

struct QueryStats final
{
    FileSystemKind fileSystemKind = FileSystemKind::Unsupported;
    bool snapshotLoaded           = false;
    bool snapshotSaved            = false;
    bool journalAvailable         = false;
    bool journalReplayApplied     = false;
    bool rebuiltJournalIdMismatch = false;
    bool rebuiltJournalRangeInvalid = false;
    bool rebuiltSnapshotCorruption  = false;
    bool usedNtfsEnumeration        = false;
    bool usedTraversalSeed          = false;
    uint64_t entryCount             = 0u;
    uint64_t fileCount              = 0u;
    uint64_t directoryCount         = 0u;
    uint64_t candidateCount         = 0u;
    uint64_t nextUsn                = 0u;
    uint64_t journalId              = 0u;
    uint64_t snapshotFileBytes      = 0u;
    uint64_t estimatedMemoryBytes   = 0u;
    uint32_t ensureReadyDurationMs  = 0u;
    uint32_t executeQueryDurationMs = 0u;
    std::wstring snapshotPath;
};

using CancelCheckFn = HRESULT(STDMETHODCALLTYPE*)(void* cookie) noexcept;

#ifdef _DEBUG
enum class SnapshotCorruptionMode : uint32_t
{
    InvalidMagic,
    JournalIdMismatch,
    NextUsnPastEnd,
};
#endif

class Repository final
{
public:
    explicit Repository(RepositoryOptions options = {}) noexcept;
    ~Repository()                           = default;
    Repository(const Repository&)           = delete;
    Repository(Repository&&)                = delete;
    Repository& operator=(const Repository&) = delete;
    Repository& operator=(Repository&&)     = delete;

    HRESULT ProbePath(std::wstring_view rootPath, SupportInfo& outSupport) noexcept;
    HRESULT Query(const QueryPlan& plan,
                  CancelCheckFn cancelCheck,
                  void* cancelCookie,
                  std::vector<Candidate>& outCandidates,
                  QueryStats* outStats) noexcept;
    HRESULT InvalidateRoot(std::wstring_view rootPath, bool deleteSnapshot) noexcept;

#ifdef _DEBUG
    HRESULT DropCachedVolumeForTests(std::wstring_view rootPath) noexcept;
    HRESULT CorruptSnapshotForTests(std::wstring_view rootPath, SnapshotCorruptionMode mode) noexcept;
#endif

private:
    RepositoryOptions _options;
    std::mutex _mutex;
    std::unordered_map<std::wstring, std::shared_ptr<VolumeIndex>> _volumes;
};
} // namespace LocalSearchIndexCore
