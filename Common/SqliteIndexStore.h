#pragma once

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include <windows.h>

#include "LocalSearchIndexCore.h"

namespace SqliteIndexStore
{
constexpr uint32_t kSchemaVersion                 = 2u;
constexpr uint32_t kMinimumWalSqliteVersionNumber = 3051002u; // 3.51.2

struct StoreInfo final
{
    std::wstring databasePath;
    uint64_t databaseBytes = 0u;
    std::wstring writeAheadLogPath;
    uint64_t writeAheadLogBytes      = 0u;
    uint32_t schemaVersion           = 0u;
    uint64_t volumeCount             = 0u;
    uint64_t entryCount              = 0u;
    uint64_t legacyImportVolumeCount = 0u;
    uint64_t pageCount               = 0u;
    uint64_t freelistPageCount       = 0u;
    std::wstring lastCheckpointUtc;
    std::wstring lastCompactionUtc;
    bool walEnabled                   = false;
    bool foreignKeysEnabled           = false;
    bool incrementalAutoVacuumEnabled = false;
    bool schemaReady                  = false;
};

constexpr uint32_t kVolumeStateReady                  = 1u;
constexpr uint32_t kVolumeStateImportedLegacySnapshot = 2u;
constexpr uint32_t kVolumeStateCurrentnessUnproven    = 3u;

struct ImportedEntry final
{
    uint64_t fileIdLow    = 0u;
    uint64_t fileIdHigh   = 0u;
    uint64_t parentIdLow  = 0u;
    uint64_t parentIdHigh = 0u;
    std::wstring fullPath;
    std::wstring name;
    unsigned long attributes     = 0u;
    uint64_t sizeBytes           = 0u;
    uint64_t writeTime100ns      = 0u;
    uint64_t creationTime100ns   = 0u;
    uint64_t lastAccessTime100ns = 0u;
    uint64_t changeTime100ns     = 0u;
    uint64_t allocationSize      = 0u;
};

struct ReplaceVolumeRequest final
{
    std::wstring rootPath;
    LocalSearchIndexCore::FileSystemKind fileSystemKind = LocalSearchIndexCore::FileSystemKind::Unsupported;
    uint64_t journalId                                  = 0u;
    uint64_t nextUsn                                    = 0u;
    uint32_t state                                      = kVolumeStateReady;
    std::vector<ImportedEntry> entries;
};

struct ReplaceVolumeResult final
{
    bool insertedNewVolume      = false;
    uint64_t importedEntryCount = 0u;
};

struct DeletedEntryId final
{
    uint64_t fileIdLow  = 0u;
    uint64_t fileIdHigh = 0u;
};

struct ApplyJournalDeltaRequest final
{
    std::wstring rootPath;
    LocalSearchIndexCore::FileSystemKind fileSystemKind = LocalSearchIndexCore::FileSystemKind::Unsupported;
    uint64_t journalId                                  = 0u;
    uint64_t nextUsn                                    = 0u;
    uint32_t state                                      = kVolumeStateReady;
    uint32_t seedStateIfMissing                         = kVolumeStateReady;
    std::vector<DeletedEntryId> deletedEntries;
    std::vector<ImportedEntry> upsertEntries;
    std::vector<ImportedEntry> seedEntriesIfMissing;
};

struct ApplyJournalDeltaResult final
{
    bool insertedNewVolume      = false;
    uint64_t deletedEntryCount  = 0u;
    uint64_t upsertedEntryCount = 0u;
};

struct QueryRequest final
{
    std::wstring rootPath;
    std::wstring namePattern;
    FileSystemSearchNameMode nameMode = FILESYSTEM_SEARCH_NAME_DISABLED;
    bool matchCaseName                = false;
    bool recursive                    = false;
    bool includeFiles                 = false;
    bool includeDirectories           = false;
    uint64_t maxResults               = 0u;
};

struct VolumeInfo final
{
    LocalSearchIndexCore::FileSystemKind fileSystemKind = LocalSearchIndexCore::FileSystemKind::Unsupported;
    uint64_t journalId                                  = 0u;
    uint64_t nextUsn                                    = 0u;
    uint64_t entryCount                                 = 0u;
    uint32_t state                                      = 0u;
};

struct QueryRuntimeStats final
{
    uint64_t scannedRows                                = 0u;
    uint64_t emittedRows                                = 0u;
    LocalSearchIndexCore::FileSystemKind fileSystemKind = LocalSearchIndexCore::FileSystemKind::Unsupported;
    uint64_t entryCount                                 = 0u;
    uint64_t fileCount                                  = 0u;
    uint64_t directoryCount                             = 0u;
    uint64_t journalId                                  = 0u;
    uint64_t nextUsn                                    = 0u;
    bool volumeReady                                    = false;
    bool readOnlyConnection                             = false;
    bool usedNamePrefilter                              = false;
};

struct ManualMaintenanceResult final
{
    StoreInfo before{};
    StoreInfo after{};
    bool ranVacuum = false;
};

struct AutomaticMaintenanceResult final
{
    StoreInfo before{};
    StoreInfo after{};
    bool maintenanceNeeded        = false;
    bool ranCheckpoint            = false;
    bool ranIncrementalVacuum     = false;
    uint64_t requestedVacuumPages = 0u;
    uint64_t reclaimedVacuumPages = 0u;
};

[[nodiscard]] HRESULT EnsureBootstrap(std::wstring_view databasePath, StoreInfo* outInfo) noexcept;
[[nodiscard]] HRESULT InspectStore(std::wstring_view databasePath, StoreInfo& outInfo) noexcept;
[[nodiscard]] HRESULT InspectVolume(std::wstring_view databasePath, std::wstring_view rootPath, VolumeInfo& outInfo) noexcept;
[[nodiscard]] HRESULT LoadVolume(std::wstring_view databasePath, std::wstring_view rootPath, ReplaceVolumeRequest& outVolume) noexcept;
[[nodiscard]] HRESULT ReplaceVolume(std::wstring_view databasePath, const ReplaceVolumeRequest& request, ReplaceVolumeResult* outResult) noexcept;
[[nodiscard]] HRESULT ApplyJournalDelta(std::wstring_view databasePath, const ApplyJournalDeltaRequest& request, ApplyJournalDeltaResult* outResult) noexcept;
[[nodiscard]] HRESULT DeleteVolume(std::wstring_view databasePath, std::wstring_view rootPath) noexcept;
[[nodiscard]] HRESULT RunManualMaintenance(std::wstring_view databasePath, ManualMaintenanceResult* outResult) noexcept;
[[nodiscard]] HRESULT RunAutomaticMaintenance(std::wstring_view databasePath,
                                              const LocalSearchIndexCore::SqliteMaintenancePolicy& policy,
                                              AutomaticMaintenanceResult* outResult) noexcept;
[[nodiscard]] HRESULT EnumerateVolume(std::wstring_view databasePath,
                                      const QueryRequest& request,
                                      LocalSearchIndexCore::CancelCheckFn cancelCheck,
                                      void* cancelCookie,
                                      LocalSearchIndexCore::CandidateCallbackFn candidateCallback,
                                      void* candidateCookie,
                                      QueryRuntimeStats* outStats) noexcept;
} // namespace SqliteIndexStore
