#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <cstdint>
#include <filesystem>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#pragma warning(pop)

#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"

namespace ViewerSqliteEngine
{
constexpr uint32_t kDefaultPageSize    = 200u;
constexpr uint32_t kDefaultQueryRowCap = 2000u;

struct Config
{
    uint32_t pageSize         = kDefaultPageSize;
    uint32_t queryRowCap      = kDefaultQueryRowCap;
    bool directOpenLocalFiles = true;
};

struct TableInfo
{
    std::wstring name;
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

class DatabaseSource final
{
public:
    static SourceOpenResult OpenFromPath(std::filesystem::path localPath, std::wstring displayName, bool deleteOnClose) noexcept;

    DatabaseSource(const DatabaseSource&)            = delete;
    DatabaseSource(DatabaseSource&&)                 = delete;
    DatabaseSource& operator=(const DatabaseSource&) = delete;
    DatabaseSource& operator=(DatabaseSource&&)      = delete;

    ~DatabaseSource() noexcept;

    [[nodiscard]] const std::filesystem::path& GetLocalPath() const noexcept;
    [[nodiscard]] const std::wstring& GetDisplayName() const noexcept;

    HRESULT ListTables(std::vector<TableInfo>& tables, std::wstring& errorText) const noexcept;
    QueryPageResult LoadTablePage(std::wstring_view tableName,
                                  uint32_t pageSize,
                                  uint64_t rowOffset,
                                  size_t orderByColumnIndex        = kNoSortColumn,
                                  TableSortDirection sortDirection = TableSortDirection::None) const noexcept;
    QueryPageResult ExecuteReadOnlyQuery(std::wstring_view sql, uint32_t rowCap) const noexcept;
    ValidationResult ValidateReadOnlyQuery(std::wstring_view sql) const noexcept;

private:
    DatabaseSource(std::filesystem::path localPath, std::wstring displayName, bool deleteOnClose) noexcept;

    std::filesystem::path _localPath;
    std::wstring _displayName;
    bool _deleteOnClose = false;
};

SourceOpenResult OpenFromViewerContext(IFileSystem* fileSystem, std::wstring_view path, bool directOpenLocalFiles) noexcept;

[[nodiscard]] std::wstring QuoteIdentifier(std::wstring_view name);
[[nodiscard]] std::wstring BuildTablePreviewSql(std::wstring_view tableName,
                                                uint32_t pageSize,
                                                uint64_t rowOffset,
                                                size_t orderByColumnIndex        = kNoSortColumn,
                                                TableSortDirection sortDirection = TableSortDirection::None);

} // namespace ViewerSqliteEngine
