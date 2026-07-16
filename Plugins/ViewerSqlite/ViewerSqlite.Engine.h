#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <cstdint>
#include <filesystem>
#include <memory>
#include <mutex>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"

struct sqlite3;

namespace ViewerSqliteEngine
{
constexpr uint32_t kDefaultPageSize    = 200u;
constexpr uint32_t kDefaultQueryRowCap = 2000u;
constexpr uint64_t kMaxSnapshotBytes   = 8ull * 1024ull * 1024ull * 1024ull;

struct QueryCancellation final
{
    const std::atomic_uint64_t* generation = nullptr;
    uint64_t expectedGeneration            = 0;

    [[nodiscard]] bool IsCancellationRequested() const noexcept
    {
        return generation != nullptr && generation->load(std::memory_order_relaxed) != expectedGeneration;
    }
};

struct QueryWorkBudget final
{
    uint64_t maxVmSteps       = 50'000'000u;
    uint32_t maxElapsedMs     = 5'000u;
    uint64_t maxResultCells   = 500'000u;
    uint64_t maxResultChars   = 16u * 1024u * 1024u;
};

struct Config
{
    uint32_t pageSize         = kDefaultPageSize;
    uint32_t queryRowCap      = kDefaultQueryRowCap;
    bool directOpenLocalFiles = true;
};

struct TableInfo
{
    // `name` is the exact SQLite identifier and is never rendered directly.
    std::wstring name;
    std::wstring displayName;
    std::wstring kind;
};

enum class TableSortDirection : uint8_t
{
    None,
    Ascending,
    Descending,
};

constexpr size_t kNoSortColumn = static_cast<size_t>(-1);

struct ColumnInfo
{
    std::wstring name;
    std::wstring declaredType;
};

struct QueryPage
{
    std::wstring executedSql;
    std::vector<ColumnInfo> columns;
    std::vector<std::vector<std::wstring>> rows;
    uint64_t rowOffset = 0;
    bool hasMore       = false;
    bool truncated     = false;
};

class DatabaseSource;

struct SourceOpenResult
{
    HRESULT hr = E_FAIL;
    std::wstring errorText;
    std::shared_ptr<DatabaseSource> source;
    std::vector<TableInfo> tables;
};

struct QueryPageResult
{
    HRESULT hr = E_FAIL;
    std::wstring errorText;
    QueryPage page;
};

struct ValidationResult
{
    HRESULT hr    = E_FAIL;
    bool accepted = false;
    std::wstring errorText;
};

enum class SnapshotKind : uint8_t
{
    LocalSqliteBackup,
    VirtualByteCopy,
};

struct SourceDebugSnapshot final
{
    uint64_t cachedConnectionOpenCount = 0;
    uint64_t operationCount             = 0;
    uint64_t cancelledOperationCount    = 0;
    uint64_t workLimitFailureCount      = 0;
    uint32_t maxConcurrentConnectionUse = 0;
    uint64_t snapshotBytes              = 0;
    SnapshotKind snapshotKind           = SnapshotKind::LocalSqliteBackup;
};

struct SqliteConnectionDeleter final
{
    void operator()(::sqlite3* db) const noexcept;
};

using unique_sqlite3 = std::unique_ptr<::sqlite3, SqliteConnectionDeleter>;

class DatabaseSource final
{
public:
    // Always snapshots the local database through SQLite's backup API. The returned source never
    // exposes subsequent changes to the live database or its WAL.
    static SourceOpenResult OpenFromPath(std::filesystem::path localPath,
                                         std::wstring displayName,
                                         QueryCancellation cancellation = {},
                                         uint64_t maxSnapshotBytes       = kMaxSnapshotBytes) noexcept;

    DatabaseSource(const DatabaseSource&)            = delete;
    DatabaseSource(DatabaseSource&&)                 = delete;
    DatabaseSource& operator=(const DatabaseSource&) = delete;
    DatabaseSource& operator=(DatabaseSource&&)      = delete;

    ~DatabaseSource() noexcept;

    [[nodiscard]] const std::filesystem::path& GetLocalPath() const noexcept;
    [[nodiscard]] const std::wstring& GetDisplayName() const noexcept;
    [[nodiscard]] SourceDebugSnapshot GetDebugSnapshot() const noexcept;

    HRESULT ListTables(std::vector<TableInfo>& tables,
                       std::wstring& errorText,
                       QueryCancellation cancellation = {},
                       QueryWorkBudget budget          = {}) const noexcept;
    QueryPageResult LoadTablePage(std::wstring_view tableName,
                                  uint32_t pageSize,
                                  uint64_t rowOffset,
                                  size_t orderByColumnIndex        = kNoSortColumn,
                                  TableSortDirection sortDirection = TableSortDirection::None,
                                  QueryCancellation cancellation   = {},
                                  QueryWorkBudget budget            = {}) const noexcept;
    QueryPageResult ExecuteReadOnlyQuery(std::wstring_view sql,
                                         uint32_t rowCap,
                                         QueryCancellation cancellation = {},
                                         QueryWorkBudget budget          = {}) const noexcept;
    ValidationResult ValidateReadOnlyQuery(std::wstring_view sql,
                                           QueryCancellation cancellation = {},
                                           QueryWorkBudget budget          = {}) const noexcept;

private:
    DatabaseSource(std::filesystem::path localPath,
                   std::wstring displayName,
                   wil::unique_handle snapshotLifetimeHandle,
                   unique_sqlite3 connection,
                   SnapshotKind snapshotKind,
                   uint64_t snapshotBytes) noexcept;

    static SourceOpenResult OpenOwnedSnapshot(std::filesystem::path snapshotPath,
                                              std::wstring displayName,
                                              wil::unique_handle snapshotLifetimeHandle,
                                              SnapshotKind snapshotKind,
                                              uint64_t snapshotBytes,
                                              QueryCancellation cancellation) noexcept;

    HRESULT ListTablesLocked(std::vector<TableInfo>& tables,
                             std::wstring& errorText,
                             QueryCancellation cancellation,
                             QueryWorkBudget budget) const noexcept;

    friend SourceOpenResult OpenFromViewerContext(
        IFileSystem* fileSystem, std::wstring_view path, bool directOpenLocalFiles, QueryCancellation cancellation) noexcept;

    std::filesystem::path _localPath;
    std::wstring _displayName;
    wil::unique_handle _snapshotLifetimeHandle;
    unique_sqlite3 _connection;
    SnapshotKind _snapshotKind = SnapshotKind::LocalSqliteBackup;
    uint64_t _snapshotBytes    = 0;

    mutable std::mutex _connectionMutex;
    mutable std::unordered_map<std::wstring, size_t> _tableColumnCounts;
    mutable std::atomic_uint64_t _operationCount{0};
    mutable std::atomic_uint64_t _cancelledOperationCount{0};
    mutable std::atomic_uint64_t _workLimitFailureCount{0};
    mutable std::atomic_uint32_t _activeConnectionUse{0};
    mutable std::atomic_uint32_t _maxConcurrentConnectionUse{0};
};

SourceOpenResult OpenFromViewerContext(IFileSystem* fileSystem,
                                       std::wstring_view path,
                                       bool directOpenLocalFiles,
                                       QueryCancellation cancellation = {}) noexcept;

[[nodiscard]] std::wstring QuoteIdentifier(std::wstring_view name);
[[nodiscard]] std::wstring BuildTablePreviewSql(std::wstring_view tableName,
                                                uint32_t pageSize,
                                                uint64_t rowOffset,
                                                size_t orderByColumnIndex        = kNoSortColumn,
                                                TableSortDirection sortDirection = TableSortDirection::None);

} // namespace ViewerSqliteEngine
