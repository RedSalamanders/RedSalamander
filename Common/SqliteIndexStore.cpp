#include "SqliteIndexStore.h"

#include "Helpers.h"

#include <algorithm>
#include <cstdlib>
#include <cstring>
#include <filesystem>
#include <format>
#include <memory>
#include <span>
#include <string_view>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <sqlite3.h>
#pragma warning(pop)

namespace SqliteIndexStore
{
namespace
{
struct ConnectionCloser final
{
    void operator()(sqlite3* db) const noexcept
    {
        if (db != nullptr)
        {
            static_cast<void>(sqlite3_close(db));
        }
    }
};

struct StatementCloser final
{
    void operator()(sqlite3_stmt* statement) const noexcept
    {
        if (statement != nullptr)
        {
            static_cast<void>(sqlite3_finalize(statement));
        }
    }
};

using unique_connection = std::unique_ptr<sqlite3, ConnectionCloser>;
using unique_statement  = std::unique_ptr<sqlite3_stmt, StatementCloser>;

constexpr int kSqliteAutoVacuumIncremental             = 2;
constexpr int kSqliteBusyTimeoutMs                     = 5'000;
constexpr uint64_t kQueryCancelCheckInterval           = 256u;
constexpr uint64_t kAutomaticIncrementalVacuumMaxPages = 4'096u;
constexpr std::string_view kMetaLastCheckpointUtc      = "last_checkpoint_utc";
constexpr std::string_view kMetaLastCompactionUtc      = "last_compaction_utc";
#ifdef ENABLE_TESTS
constexpr wchar_t kEnumerateFailAfterEmittedRowsEnvVar[] = L"REDSALAMANDER_TEST_SQLITE_FAIL_AFTER_EMITTED_ROWS";
#endif

[[nodiscard]] HRESULT BindWideText(sqlite3_stmt* statement, int index, std::wstring_view value);

[[nodiscard]] std::wstring FormatSqliteVersion(const uint32_t versionNumber)
{
    const uint32_t major = versionNumber / 1'000'000u;
    const uint32_t minor = (versionNumber / 1'000u) % 1'000u;
    const uint32_t patch = versionNumber % 1'000u;
    return std::format(L"{}.{}.{}", major, minor, patch);
}

[[nodiscard]] std::wstring GetCurrentUtcTimestampText() noexcept
{
    SYSTEMTIME now{};
    ::GetSystemTime(&now);
    return std::format(L"{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
                       static_cast<unsigned int>(now.wYear),
                       static_cast<unsigned int>(now.wMonth),
                       static_cast<unsigned int>(now.wDay),
                       static_cast<unsigned int>(now.wHour),
                       static_cast<unsigned int>(now.wMinute),
                       static_cast<unsigned int>(now.wSecond));
}

[[nodiscard]] std::wstring NormalizeDatabasePath(std::wstring_view databasePath)
{
    if (databasePath.empty())
    {
        return {};
    }

    return std::filesystem::path(databasePath).lexically_normal().wstring();
}

#ifdef ENABLE_TESTS
[[nodiscard]] uint64_t GetInjectedEnumerateFailureAfterEmittedRows() noexcept
{
    const DWORD required = ::GetEnvironmentVariableW(kEnumerateFailAfterEmittedRowsEnvVar, nullptr, 0u);
    if (required <= 1u)
    {
        return 0u;
    }

    std::wstring value;
    value.resize(required - 1u);
    const DWORD written = ::GetEnvironmentVariableW(kEnumerateFailAfterEmittedRowsEnvVar, value.data(), required);
    if (written == 0u || written >= required)
    {
        return 0u;
    }

    value.resize(written);
    wchar_t* end                    = nullptr;
    const unsigned long long parsed = _wcstoui64(value.c_str(), &end, 10);
    if (end == nullptr || *end != L'\0')
    {
        return 0u;
    }

    return static_cast<uint64_t>(parsed);
}
#endif

[[nodiscard]] std::string ToUtf8(std::wstring_view text)
{
    if (text.empty())
    {
        return {};
    }

    const std::u8string utf8 = std::filesystem::path(text).u8string();
    return {reinterpret_cast<const char*>(utf8.data()), utf8.size()};
}

[[nodiscard]] std::wstring FoldText(std::wstring_view text) noexcept
{
    return OrdinalString::FoldCaseInvariant(text);
}

[[nodiscard]] std::wstring GetExtensionFolded(std::wstring_view name) noexcept
{
    const std::wstring extension = std::filesystem::path(name).extension().wstring();
    return FoldText(extension);
}

[[nodiscard]] std::wstring GetSqliteErrorMessage(sqlite3* db)
{
    if (db == nullptr)
    {
        return {};
    }

    const auto* message = static_cast<const wchar_t*>(sqlite3_errmsg16(db));
    return (message != nullptr) ? std::wstring(message) : std::wstring{};
}

[[nodiscard]] uint64_t GetFileBytes(std::wstring_view path)
{
    if (path.empty())
    {
        return 0u;
    }

    std::error_code ec;
    const uintmax_t size = std::filesystem::file_size(std::filesystem::path(path), ec);
    if (ec)
    {
        return 0u;
    }

    return static_cast<uint64_t>(size);
}

[[nodiscard]] HRESULT EnsureParentDirectory(std::wstring_view databasePath)
{
    const std::filesystem::path path(databasePath);
    const std::filesystem::path parent = path.parent_path();
    if (parent.empty())
    {
        return S_OK;
    }

    std::error_code ec;
    if (std::filesystem::create_directories(parent, ec) || ! ec)
    {
        return S_OK;
    }

    const std::string ecMessage = ec.message();
    Debug::Error(L"SqliteIndexStore: failed to create parent directory '{}'. ec={} message='{}'",
                 parent.wstring(),
                 ec.value(),
                 std::wstring(ecMessage.begin(), ecMessage.end()));
    return HRESULT_FROM_WIN32(static_cast<unsigned long>(ec.value()));
}

[[nodiscard]] HRESULT OpenConnection(std::wstring_view databasePath, int flags, unique_connection& outConnection)
{
    sqlite3* raw               = nullptr;
    const std::string utf8Path = ToUtf8(databasePath);
    const int sqliteResult     = sqlite3_open_v2(utf8Path.c_str(), &raw, flags, nullptr);
    unique_connection connection(raw);
    if (sqliteResult != SQLITE_OK)
    {
        Debug::Error(
            L"SqliteIndexStore: sqlite3_open_v2 failed for '{}'. code={} message='{}'", std::wstring(databasePath), sqliteResult, GetSqliteErrorMessage(raw));
        return E_FAIL;
    }

    static_cast<void>(sqlite3_busy_timeout(connection.get(), kSqliteBusyTimeoutMs));
    outConnection = std::move(connection);
    return S_OK;
}

[[nodiscard]] HRESULT ExecuteSql(sqlite3* db, std::string_view sql, std::wstring_view context)
{
    const std::string sqlText(sql);
    char* errorText        = nullptr;
    const int sqliteResult = sqlite3_exec(db, sqlText.c_str(), nullptr, nullptr, &errorText);
    wil::unique_hlocal_string sqliteError{};
    if (errorText != nullptr)
    {
        const size_t errorLen = std::strlen(errorText);
        const int required    = ::MultiByteToWideChar(CP_UTF8, 0, errorText, static_cast<int>(errorLen), nullptr, 0);
        if (required > 0)
        {
            sqliteError.reset(static_cast<wchar_t*>(::LocalAlloc(LMEM_FIXED, static_cast<size_t>(required + 1) * sizeof(wchar_t))));
            if (sqliteError)
            {
                static_cast<void>(::MultiByteToWideChar(CP_UTF8, 0, errorText, static_cast<int>(errorLen), sqliteError.get(), required));
                sqliteError.get()[required] = L'\0';
            }
        }
        sqlite3_free(errorText);
    }

    if (sqliteResult != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: {} failed. code={} message='{}'", context, sqliteResult, sqliteError ? sqliteError.get() : GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    return S_OK;
}

[[nodiscard]] HRESULT PrepareStatement(sqlite3* db, std::string_view sql, unique_statement& outStatement)
{
    sqlite3_stmt* raw      = nullptr;
    const int sqliteResult = sqlite3_prepare_v2(db, sql.data(), static_cast<int>(sql.size()), &raw, nullptr);
    if (sqliteResult != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: sqlite3_prepare_v2 failed. code={} message='{}'", sqliteResult, GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    outStatement.reset(raw);
    return S_OK;
}

[[nodiscard]] HRESULT ReadSingleInt(sqlite3* db, std::string_view sql, int& outValue)
{
    unique_statement statement;
    HRESULT hr = PrepareStatement(db, sql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult != SQLITE_ROW)
    {
        Debug::Error(L"SqliteIndexStore: expected one-row integer result. sql='{}' code={} message='{}'",
                     std::wstring(sql.begin(), sql.end()),
                     sqliteResult,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    outValue = sqlite3_column_int(statement.get(), 0);
    return S_OK;
}

[[nodiscard]] HRESULT ReadSingleInt64(sqlite3* db, std::string_view sql, uint64_t& outValue)
{
    unique_statement statement;
    HRESULT hr = PrepareStatement(db, sql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult != SQLITE_ROW)
    {
        Debug::Error(L"SqliteIndexStore: expected one-row int64 result. sql='{}' code={} message='{}'",
                     std::wstring(sql.begin(), sql.end()),
                     sqliteResult,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    outValue = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 0));
    return S_OK;
}

[[nodiscard]] HRESULT ReadSingleText(sqlite3* db, std::string_view sql, std::wstring& outValue)
{
    unique_statement statement;
    HRESULT hr = PrepareStatement(db, sql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult != SQLITE_ROW)
    {
        Debug::Error(L"SqliteIndexStore: expected one-row text result. sql='{}' code={} message='{}'",
                     std::wstring(sql.begin(), sql.end()),
                     sqliteResult,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const auto* text = static_cast<const wchar_t*>(sqlite3_column_text16(statement.get(), 0));
    outValue         = (text != nullptr) ? std::wstring(text) : std::wstring{};
    return S_OK;
}

[[nodiscard]] HRESULT ReadMetaText(sqlite3* db, std::string_view key, std::wstring& outValue, bool& outFound)
{
    outValue.clear();
    outFound = false;

    unique_statement statement;
    constexpr std::string_view kSql = "SELECT value FROM meta WHERE key = ?1;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int bindResult = sqlite3_bind_text(statement.get(), 1, key.data(), static_cast<int>(key.size()), SQLITE_TRANSIENT);
    if (bindResult != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: failed to bind meta key '{}'. code={} message='{}'",
                     std::wstring(key.begin(), key.end()),
                     bindResult,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult == SQLITE_DONE)
    {
        return S_OK;
    }
    if (sqliteResult != SQLITE_ROW)
    {
        Debug::Error(L"SqliteIndexStore: meta lookup failed for key '{}'. code={} message='{}'",
                     std::wstring(key.begin(), key.end()),
                     sqliteResult,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const auto* text = static_cast<const wchar_t*>(sqlite3_column_text16(statement.get(), 0));
    outValue         = (text != nullptr) ? std::wstring(text) : std::wstring{};
    outFound         = true;
    return S_OK;
}

[[nodiscard]] HRESULT UpsertMetaText(sqlite3* db, std::string_view key, std::wstring_view value)
{
    unique_statement statement;
    constexpr std::string_view kSql = "INSERT INTO meta(key, value) VALUES(?1, ?2) "
                                      "ON CONFLICT(key) DO UPDATE SET value = excluded.value;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int bindKey       = sqlite3_bind_text(statement.get(), 1, key.data(), static_cast<int>(key.size()), SQLITE_TRANSIENT);
    const HRESULT bindValue = BindWideText(statement.get(), 2, value);
    if (bindKey != SQLITE_OK || FAILED(bindValue))
    {
        Debug::Error(L"SqliteIndexStore: failed to bind meta upsert for key '{}'. code={} message='{}'",
                     std::wstring(key.begin(), key.end()),
                     bindKey,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult != SQLITE_DONE)
    {
        Debug::Error(L"SqliteIndexStore: meta upsert failed for key '{}'. code={} message='{}'",
                     std::wstring(key.begin(), key.end()),
                     sqliteResult,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    return S_OK;
}

[[nodiscard]] HRESULT SchemaObjectExists(sqlite3* db, std::string_view type, std::string_view name, bool& outExists)
{
    unique_statement statement;
    constexpr std::string_view kSql = "SELECT 1 FROM sqlite_master WHERE type = ?1 AND name = ?2 LIMIT 1;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int bindType = sqlite3_bind_text(statement.get(), 1, type.data(), static_cast<int>(type.size()), SQLITE_TRANSIENT);
    const int bindName = sqlite3_bind_text(statement.get(), 2, name.data(), static_cast<int>(name.size()), SQLITE_TRANSIENT);
    if (bindType != SQLITE_OK || bindName != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: failed to bind schema-object lookup. type='{}' name='{}' code={} message='{}'",
                     std::wstring(type.begin(), type.end()),
                     std::wstring(name.begin(), name.end()),
                     (bindType != SQLITE_OK) ? bindType : bindName,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult != SQLITE_ROW && sqliteResult != SQLITE_DONE)
    {
        Debug::Error(L"SqliteIndexStore: schema-object lookup failed. type='{}' name='{}' code={} message='{}'",
                     std::wstring(type.begin(), type.end()),
                     std::wstring(name.begin(), name.end()),
                     sqliteResult,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    outExists = (sqliteResult == SQLITE_ROW);
    return S_OK;
}

[[nodiscard]] HRESULT TableColumnExists(sqlite3* db, std::string_view tableName, std::string_view columnName, bool& outExists)
{
    outExists = false;

    unique_statement statement;
    const std::string sql = std::format("PRAGMA table_info({});", tableName);
    HRESULT hr            = PrepareStatement(db, sql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    while (true)
    {
        const int sqliteResult = sqlite3_step(statement.get());
        if (sqliteResult == SQLITE_DONE)
        {
            return S_OK;
        }
        if (sqliteResult != SQLITE_ROW)
        {
            Debug::Error(L"SqliteIndexStore: table_info query failed. table='{}' code={} message='{}'",
                         std::wstring(tableName.begin(), tableName.end()),
                         sqliteResult,
                         GetSqliteErrorMessage(db));
            return E_FAIL;
        }

        const auto* text = reinterpret_cast<const char*>(sqlite3_column_text(statement.get(), 1));
        if (text != nullptr && columnName == std::string_view(text))
        {
            outExists = true;
            return S_OK;
        }
    }
}

[[nodiscard]] HRESULT ReadPageCounts(sqlite3* db, uint64_t& outPageCount, uint64_t& outFreelistPageCount)
{
    outPageCount         = 0u;
    outFreelistPageCount = 0u;

    HRESULT hr = ReadSingleInt64(db, "PRAGMA page_count;", outPageCount);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReadSingleInt64(db, "PRAGMA freelist_count;", outFreelistPageCount);
    if (FAILED(hr))
    {
        return hr;
    }

    return S_OK;
}

[[nodiscard]] HRESULT ReadPageSize(sqlite3* db, uint64_t& outPageSize)
{
    outPageSize = 0u;

    int pageSize     = 0;
    const HRESULT hr = ReadSingleInt(db, "PRAGMA page_size;", pageSize);
    if (FAILED(hr))
    {
        return hr;
    }

    if (pageSize <= 0)
    {
        return E_FAIL;
    }

    outPageSize = static_cast<uint64_t>(pageSize);
    return S_OK;
}

[[nodiscard]] HRESULT BindWideText(sqlite3_stmt* statement, int index, std::wstring_view value)
{
    const int sqliteResult = sqlite3_bind_text16(statement, index, value.data(), static_cast<int>(value.size() * sizeof(wchar_t)), SQLITE_TRANSIENT);
    if (sqliteResult != SQLITE_OK)
    {
        return E_FAIL;
    }

    return S_OK;
}

[[nodiscard]] HRESULT RunWalCheckpoint(sqlite3* db, const int mode, std::wstring_view context)
{
    int logFrames          = 0;
    int checkpointedFrames = 0;
    const int sqliteResult = sqlite3_wal_checkpoint_v2(db, nullptr, mode, &logFrames, &checkpointedFrames);
    if (sqliteResult != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: {} failed. code={} logFrames={} checkpointedFrames={} message='{}'",
                     context,
                     sqliteResult,
                     logFrames,
                     checkpointedFrames,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    return S_OK;
}

[[nodiscard]] uint64_t CalculateFragmentationPercent(const StoreInfo& info) noexcept
{
    if (info.pageCount == 0u || info.freelistPageCount == 0u)
    {
        return 0u;
    }

    return (info.freelistPageCount * 100u) / info.pageCount;
}

[[nodiscard]] uint64_t CalculateReclaimableBytes(const StoreInfo& info, const uint64_t pageSizeBytes) noexcept
{
    if (pageSizeBytes == 0u || info.freelistPageCount == 0u)
    {
        return 0u;
    }

    return info.freelistPageCount * pageSizeBytes;
}

struct VolumeRow final
{
    sqlite3_int64 volumeId                              = 0;
    LocalSearchIndexCore::FileSystemKind fileSystemKind = LocalSearchIndexCore::FileSystemKind::Unsupported;
    uint64_t journalId                                  = 0u;
    uint64_t nextUsn                                    = 0u;
    uint64_t entryCount                                 = 0u;
    uint32_t state                                      = 0u;
};

struct RootEntryRow final
{
    uint64_t fileIdLow       = 0u;
    uint64_t fileIdHigh      = 0u;
    unsigned long attributes = 0u;
    std::wstring fullPath;
    std::wstring name;
};

enum class NamePrefilterKind : uint32_t
{
    None = 0,
    ExactFolded,
    PrefixFolded,
    ExtensionFolded,
};

struct NamePrefilter final
{
    NamePrefilterKind kind = NamePrefilterKind::None;
    std::wstring value;
};

[[nodiscard]] bool ContainsWildcardTokens(std::wstring_view pattern) noexcept
{
    return pattern.find_first_of(L"*?") != std::wstring_view::npos;
}

[[nodiscard]] bool TryBuildNamePrefilter(const QueryRequest& request, NamePrefilter& outPrefilter) noexcept
{
    outPrefilter = {};
    if (request.matchCaseName)
    {
        return false;
    }

    switch (request.nameMode)
    {
        case FILESYSTEM_SEARCH_NAME_DISABLED: return false;
        case FILESYSTEM_SEARCH_NAME_LITERAL:
            outPrefilter.kind  = NamePrefilterKind::ExactFolded;
            outPrefilter.value = FoldText(request.namePattern);
            return ! outPrefilter.value.empty();
        case FILESYSTEM_SEARCH_NAME_WILDCARD:
            if (! ContainsWildcardTokens(request.namePattern))
            {
                outPrefilter.kind  = NamePrefilterKind::ExactFolded;
                outPrefilter.value = FoldText(request.namePattern);
                return ! outPrefilter.value.empty();
            }

            if (request.namePattern.size() > 2u && request.namePattern[0] == L'*' && request.namePattern[1] == L'.' &&
                request.namePattern.find_first_of(L"*?", 1u) == std::wstring_view::npos)
            {
                outPrefilter.kind  = NamePrefilterKind::ExtensionFolded;
                outPrefilter.value = FoldText(request.namePattern.substr(1u));
                return ! outPrefilter.value.empty();
            }

            if (request.namePattern.size() > 1u && request.namePattern.back() == L'*')
            {
                const std::wstring_view prefix = request.namePattern.substr(0u, request.namePattern.size() - 1u);
                if (! prefix.empty() && prefix.find_first_of(L"*?") == std::wstring_view::npos)
                {
                    outPrefilter.kind  = NamePrefilterKind::PrefixFolded;
                    outPrefilter.value = FoldText(prefix);
                    return ! outPrefilter.value.empty();
                }
            }

            return false;
        case FILESYSTEM_SEARCH_NAME_REGEX: return false;
    }

    return false;
}

[[nodiscard]] std::wstring ReadWideColumn(sqlite3_stmt* statement, const int columnIndex)
{
    const auto* text = static_cast<const wchar_t*>(sqlite3_column_text16(statement, columnIndex));
    if (text == nullptr)
    {
        return {};
    }

    const int byteCount = sqlite3_column_bytes16(statement, columnIndex);
    if (byteCount <= 0)
    {
        return {};
    }

    return std::wstring(text, static_cast<size_t>(byteCount) / sizeof(wchar_t));
}

[[nodiscard]] std::wstring GetDisplayName(std::wstring_view fullPath, std::wstring_view storedName)
{
    if (! storedName.empty())
    {
        return std::wstring(storedName);
    }

    return std::filesystem::path(fullPath).filename().wstring();
}

[[nodiscard]] HRESULT CheckCancelled(LocalSearchIndexCore::CancelCheckFn cancelCheck, void* cancelCookie) noexcept
{
    if (cancelCheck == nullptr)
    {
        return S_OK;
    }

    const HRESULT hr = cancelCheck(cancelCookie);
    return FAILED(hr) ? hr : S_OK;
}

[[nodiscard]] HRESULT ReadVolumeRow(sqlite3* db, std::wstring_view rootPath, VolumeRow& outRow)
{
    outRow = {};

    unique_statement statement;
    constexpr std::string_view kSql = "SELECT volume_id, fs_kind, journal_id, next_usn, entry_count, state FROM volumes WHERE root_path = ?1;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = BindWideText(statement.get(), 1, rootPath);
    if (FAILED(hr))
    {
        Debug::Error(L"SqliteIndexStore: failed to bind root_path for volume query '{}'.", std::wstring(rootPath));
        return hr;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult == SQLITE_DONE)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    if (sqliteResult != SQLITE_ROW)
    {
        Debug::Error(L"SqliteIndexStore: volume query failed for '{}'. code={} message='{}'", std::wstring(rootPath), sqliteResult, GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    outRow.volumeId       = sqlite3_column_int64(statement.get(), 0);
    outRow.fileSystemKind = static_cast<LocalSearchIndexCore::FileSystemKind>(sqlite3_column_int(statement.get(), 1));
    outRow.journalId      = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 2));
    outRow.nextUsn        = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 3));
    outRow.entryCount     = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 4));
    outRow.state          = static_cast<uint32_t>(sqlite3_column_int(statement.get(), 5));
    return S_OK;
}

[[nodiscard]] HRESULT ReadVolumeTypeCounts(sqlite3* db, sqlite3_int64 volumeId, uint64_t& outFileCount, uint64_t& outDirectoryCount)
{
    outFileCount      = 0u;
    outDirectoryCount = 0u;

    unique_statement statement;
    constexpr std::string_view kSql = "SELECT "
                                      "SUM(CASE WHEN is_dir = 0 THEN 1 ELSE 0 END), "
                                      "SUM(CASE WHEN is_dir = 1 THEN 1 ELSE 0 END) "
                                      "FROM entries WHERE volume_id = ?1;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int bindVolumeId = sqlite3_bind_int64(statement.get(), 1, volumeId);
    if (bindVolumeId != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: failed to bind volume_id for count query. code={} message='{}'", bindVolumeId, GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult != SQLITE_ROW)
    {
        Debug::Error(L"SqliteIndexStore: volume count query failed for volume_id={}. code={} message='{}'",
                     static_cast<long long>(volumeId),
                     sqliteResult,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    outFileCount      = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 0));
    outDirectoryCount = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 1));
    return S_OK;
}

[[nodiscard]] HRESULT ReadRootEntry(sqlite3* db, sqlite3_int64 volumeId, std::wstring_view rootPath, RootEntryRow& outEntry)
{
    outEntry = {};

    unique_statement statement;
    constexpr std::string_view kSql = "SELECT file_id_low, file_id_high, attributes, full_path, name "
                                      "FROM entries WHERE volume_id = ?1 AND full_path_folded = ?2 LIMIT 1;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring foldedRootPath = FoldText(rootPath);
    const int bindVolumeId            = sqlite3_bind_int64(statement.get(), 1, volumeId);
    const HRESULT bindRootPath        = BindWideText(statement.get(), 2, foldedRootPath);
    if (bindVolumeId != SQLITE_OK || FAILED(bindRootPath))
    {
        Debug::Error(L"SqliteIndexStore: failed to bind root entry query for '{}'. code={} message='{}'",
                     std::wstring(rootPath),
                     bindVolumeId,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult == SQLITE_DONE)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    if (sqliteResult != SQLITE_ROW)
    {
        Debug::Error(
            L"SqliteIndexStore: root entry query failed for '{}'. code={} message='{}'", std::wstring(rootPath), sqliteResult, GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    outEntry.fileIdLow  = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 0));
    outEntry.fileIdHigh = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 1));
    outEntry.attributes = static_cast<unsigned long>(sqlite3_column_int64(statement.get(), 2));
    outEntry.fullPath   = ReadWideColumn(statement.get(), 3);
    outEntry.name       = ReadWideColumn(statement.get(), 4);
    return S_OK;
}

[[nodiscard]] HRESULT PrepareEnumerateStatement(sqlite3* db,
                                                sqlite3_int64 volumeId,
                                                const RootEntryRow& rootEntry,
                                                const QueryRequest& request,
                                                const NamePrefilter& prefilter,
                                                unique_statement& outStatement)
{
    std::string sql =
        "SELECT full_path, name, attributes, size_bytes, write_time_100ns, creation_time_100ns, last_access_time_100ns, change_time_100ns, allocation_size "
        "FROM entries WHERE volume_id = ?1";
    int parameterIndex         = 2;
    const bool rootIsDirectory = (rootEntry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;

    if (rootIsDirectory)
    {
        if (request.recursive)
        {
            sql += std::format(" AND NOT (file_id_low = ?{} AND file_id_high = ?{})", parameterIndex, parameterIndex + 1);
        }
        else
        {
            sql += std::format(" AND parent_id_low = ?{} AND parent_id_high = ?{}", parameterIndex, parameterIndex + 1);
        }
        parameterIndex += 2;
    }
    else
    {
        sql += std::format(" AND file_id_low = ?{} AND file_id_high = ?{}", parameterIndex, parameterIndex + 1);
        parameterIndex += 2;
    }

    if (request.includeFiles != request.includeDirectories)
    {
        sql += std::format(" AND is_dir = ?{}", parameterIndex++);
    }

    switch (prefilter.kind)
    {
        case NamePrefilterKind::ExactFolded: sql += std::format(" AND name_folded = ?{}", parameterIndex++); break;
        case NamePrefilterKind::PrefixFolded: sql += std::format(" AND name_folded LIKE ?{}", parameterIndex++); break;
        case NamePrefilterKind::ExtensionFolded: sql += std::format(" AND extension_folded = ?{}", parameterIndex++); break;
        case NamePrefilterKind::None: break;
    }

    sql += " ORDER BY full_path_folded;";

    HRESULT hr = PrepareStatement(db, sql, outStatement);
    if (FAILED(hr))
    {
        return hr;
    }

    int bindIndex          = 1;
    const int bindVolumeId = sqlite3_bind_int64(outStatement.get(), bindIndex++, volumeId);
    const int bindRootLow  = sqlite3_bind_int64(outStatement.get(), bindIndex++, static_cast<sqlite3_int64>(rootEntry.fileIdLow));
    const int bindRootHigh = sqlite3_bind_int64(outStatement.get(), bindIndex++, static_cast<sqlite3_int64>(rootEntry.fileIdHigh));
    if (bindVolumeId != SQLITE_OK || bindRootLow != SQLITE_OK || bindRootHigh != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: failed to bind query root state for '{}'. code={} message='{}'",
                     request.rootPath,
                     (bindVolumeId != SQLITE_OK)  ? bindVolumeId
                     : (bindRootLow != SQLITE_OK) ? bindRootLow
                                                  : bindRootHigh,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    if (request.includeFiles != request.includeDirectories)
    {
        const int bindType = sqlite3_bind_int(outStatement.get(), bindIndex++, request.includeDirectories ? 1 : 0);
        if (bindType != SQLITE_OK)
        {
            Debug::Error(
                L"SqliteIndexStore: failed to bind directory filter for '{}'. code={} message='{}'", request.rootPath, bindType, GetSqliteErrorMessage(db));
            return E_FAIL;
        }
    }

    switch (prefilter.kind)
    {
        case NamePrefilterKind::ExactFolded:
        case NamePrefilterKind::ExtensionFolded:
        {
            hr = BindWideText(outStatement.get(), bindIndex++, prefilter.value);
            if (FAILED(hr))
            {
                Debug::Error(L"SqliteIndexStore: failed to bind exact prefilter for '{}'.", request.rootPath);
                return hr;
            }
            break;
        }
        case NamePrefilterKind::PrefixFolded:
        {
            const std::wstring prefixPattern = prefilter.value + L"%";
            hr                               = BindWideText(outStatement.get(), bindIndex++, prefixPattern);
            if (FAILED(hr))
            {
                Debug::Error(L"SqliteIndexStore: failed to bind prefix prefilter for '{}'.", request.rootPath);
                return hr;
            }
            break;
        }
        case NamePrefilterKind::None: break;
    }

    return S_OK;
}

[[nodiscard]] HRESULT ApplyPragmas(sqlite3* db)
{
    const uint32_t runtimeVersion = static_cast<uint32_t>(sqlite3_libversion_number());
    if (runtimeVersion < kMinimumWalSqliteVersionNumber)
    {
        Debug::Error(L"SqliteIndexStore: SQLite runtime version {} is below the accepted WAL minimum {}.",
                     FormatSqliteVersion(runtimeVersion),
                     FormatSqliteVersion(kMinimumWalSqliteVersionNumber));
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }

    const LocalSearchIndexCore::SqliteMaintenancePolicy policy = LocalSearchIndexCore::GetDefaultSqliteMaintenancePolicy();

    int autoVacuum = 0;
    HRESULT hr     = ReadSingleInt(db, "PRAGMA auto_vacuum;", autoVacuum);
    if (FAILED(hr))
    {
        return hr;
    }

    if (autoVacuum != kSqliteAutoVacuumIncremental)
    {
        hr = ExecuteSql(db, "PRAGMA auto_vacuum=INCREMENTAL;", L"PRAGMA auto_vacuum=INCREMENTAL");
        if (FAILED(hr))
        {
            return hr;
        }

        // SQLite persists auto-vacuum mode changes only after VACUUM rewrites the database.
        hr = ExecuteSql(db, "VACUUM;", L"VACUUM");
        if (FAILED(hr))
        {
            return hr;
        }
    }

    hr = ExecuteSql(db, "PRAGMA journal_mode=WAL;", L"PRAGMA journal_mode=WAL");
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ExecuteSql(db, "PRAGMA synchronous=NORMAL;", L"PRAGMA synchronous=NORMAL");
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ExecuteSql(db, "PRAGMA temp_store=MEMORY;", L"PRAGMA temp_store=MEMORY");
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ExecuteSql(db, "PRAGMA foreign_keys=ON;", L"PRAGMA foreign_keys=ON");
    if (FAILED(hr))
    {
        return hr;
    }

    int pageSize = 0;
    hr           = ReadSingleInt(db, "PRAGMA page_size;", pageSize);
    if (FAILED(hr) || pageSize <= 0)
    {
        return FAILED(hr) ? hr : E_FAIL;
    }

    const uint64_t pageCount =
        std::max<uint64_t>(1u, (policy.autoCheckpointTargetBytes + static_cast<uint64_t>(pageSize) - 1u) / static_cast<uint64_t>(pageSize));
    hr = ExecuteSql(db, std::format("PRAGMA wal_autocheckpoint={};", pageCount), L"PRAGMA wal_autocheckpoint");
    if (FAILED(hr))
    {
        return hr;
    }

    return S_OK;
}

[[nodiscard]] HRESULT EnsureSchemaVersion1(sqlite3* db)
{
    constexpr std::string_view kSchemaSql = "BEGIN IMMEDIATE;"
                                            "PRAGMA user_version=1;"
                                            "CREATE TABLE IF NOT EXISTS meta("
                                            "    key TEXT PRIMARY KEY,"
                                            "    value TEXT NOT NULL"
                                            ");"
                                            "CREATE TABLE IF NOT EXISTS volumes("
                                            "    volume_id INTEGER PRIMARY KEY,"
                                            "    root_path TEXT UNIQUE NOT NULL,"
                                            "    fs_kind INTEGER NOT NULL,"
                                            "    journal_id INTEGER NOT NULL,"
                                            "    next_usn INTEGER NOT NULL,"
                                            "    state INTEGER NOT NULL,"
                                            "    entry_count INTEGER NOT NULL,"
                                            "    last_seed_utc TEXT NULL,"
                                            "    last_replay_utc TEXT NULL,"
                                            "    last_error_hr INTEGER NOT NULL DEFAULT 0"
                                            ");"
                                            "CREATE TABLE IF NOT EXISTS entries("
                                            "    volume_id INTEGER NOT NULL,"
                                            "    file_id_low INTEGER NOT NULL,"
                                            "    file_id_high INTEGER NOT NULL,"
                                            "    parent_id_low INTEGER NOT NULL,"
                                            "    parent_id_high INTEGER NOT NULL,"
                                            "    full_path TEXT NOT NULL,"
                                            "    full_path_folded TEXT NOT NULL,"
                                            "    name TEXT NOT NULL,"
                                            "    name_folded TEXT NOT NULL,"
                                            "    extension_folded TEXT NOT NULL,"
                                            "    attributes INTEGER NOT NULL,"
                                            "    is_dir INTEGER NOT NULL,"
                                            "    size_bytes INTEGER NOT NULL,"
                                            "    write_time_100ns INTEGER NOT NULL,"
                                            "    creation_time_100ns INTEGER NOT NULL,"
                                            "    last_access_time_100ns INTEGER NOT NULL,"
                                            "    change_time_100ns INTEGER NOT NULL,"
                                            "    allocation_size INTEGER NOT NULL,"
                                            "    PRIMARY KEY(volume_id, file_id_low, file_id_high)"
                                            ") WITHOUT ROWID;"
                                            "CREATE INDEX IF NOT EXISTS idx_entries_name_folded ON entries(volume_id, name_folded);"
                                            "CREATE INDEX IF NOT EXISTS idx_entries_extension_folded ON entries(volume_id, extension_folded, is_dir);"
                                            "CREATE INDEX IF NOT EXISTS idx_entries_full_path_folded ON entries(volume_id, full_path_folded);"
                                            "CREATE INDEX IF NOT EXISTS idx_entries_parent ON entries(volume_id, parent_id_low, parent_id_high);"
                                            "INSERT INTO meta(key, value) VALUES('schema_version', '1')"
                                            "    ON CONFLICT(key) DO UPDATE SET value = excluded.value;"
                                            "INSERT INTO meta(key, value) VALUES('store_kind', 'sqlite-v2')"
                                            "    ON CONFLICT(key) DO UPDATE SET value = excluded.value;"
                                            "COMMIT;";
    return ExecuteSql(db, kSchemaSql, L"SQLite schema bootstrap");
}

[[nodiscard]] HRESULT EnsureSchemaVersion2(sqlite3* db, bool markExistingVolumesLegacy)
{
    bool hasCreationTimeColumn   = false;
    bool hasLastAccessTimeColumn = false;
    bool hasChangeTimeColumn     = false;
    bool hasAllocationSizeColumn = false;

    HRESULT hr = TableColumnExists(db, "entries", "creation_time_100ns", hasCreationTimeColumn);
    if (SUCCEEDED(hr))
    {
        hr = TableColumnExists(db, "entries", "last_access_time_100ns", hasLastAccessTimeColumn);
    }
    if (SUCCEEDED(hr))
    {
        hr = TableColumnExists(db, "entries", "change_time_100ns", hasChangeTimeColumn);
    }
    if (SUCCEEDED(hr))
    {
        hr = TableColumnExists(db, "entries", "allocation_size", hasAllocationSizeColumn);
    }
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ExecuteSql(db, "BEGIN IMMEDIATE;", L"BEGIN IMMEDIATE");
    if (FAILED(hr))
    {
        return hr;
    }

    bool commit   = false;
    auto rollback = wil::scope_exit([&]() noexcept
    {
        if (! commit)
        {
            static_cast<void>(ExecuteSql(db, "ROLLBACK;", L"ROLLBACK"));
        }
    });

    if (! hasCreationTimeColumn)
    {
        hr = ExecuteSql(
            db, "ALTER TABLE entries ADD COLUMN creation_time_100ns INTEGER NOT NULL DEFAULT 0;", L"ALTER TABLE entries ADD COLUMN creation_time_100ns");
        if (FAILED(hr))
        {
            return hr;
        }
    }
    if (! hasLastAccessTimeColumn)
    {
        hr = ExecuteSql(
            db, "ALTER TABLE entries ADD COLUMN last_access_time_100ns INTEGER NOT NULL DEFAULT 0;", L"ALTER TABLE entries ADD COLUMN last_access_time_100ns");
        if (FAILED(hr))
        {
            return hr;
        }
    }
    if (! hasChangeTimeColumn)
    {
        hr =
            ExecuteSql(db, "ALTER TABLE entries ADD COLUMN change_time_100ns INTEGER NOT NULL DEFAULT 0;", L"ALTER TABLE entries ADD COLUMN change_time_100ns");
        if (FAILED(hr))
        {
            return hr;
        }
    }
    if (! hasAllocationSizeColumn)
    {
        hr = ExecuteSql(db, "ALTER TABLE entries ADD COLUMN allocation_size INTEGER NOT NULL DEFAULT 0;", L"ALTER TABLE entries ADD COLUMN allocation_size");
        if (FAILED(hr))
        {
            return hr;
        }
    }

    hr = ExecuteSql(db,
                    std::format("UPDATE meta SET value = '2' WHERE key = 'schema_version';"
                                "INSERT INTO meta(key, value) VALUES('schema_version', '2')"
                                "    ON CONFLICT(key) DO UPDATE SET value = excluded.value;"
                                "INSERT INTO meta(key, value) VALUES('store_kind', 'sqlite-v2')"
                                "    ON CONFLICT(key) DO UPDATE SET value = excluded.value;"
                                "PRAGMA user_version=2;"),
                    L"SQLite schema v2 metadata update");
    if (FAILED(hr))
    {
        return hr;
    }

    if (markExistingVolumesLegacy)
    {
        hr = ExecuteSql(db,
                        std::format("UPDATE volumes SET state = {} WHERE state = {};",
                                    static_cast<unsigned long long>(kVolumeStateImportedLegacySnapshot),
                                    static_cast<unsigned long long>(kVolumeStateReady)),
                        L"SQLite schema v2 volume-state migration");
        if (FAILED(hr))
        {
            return hr;
        }
    }

    hr = ExecuteSql(db, "COMMIT;", L"COMMIT");
    if (FAILED(hr))
    {
        return hr;
    }

    commit = true;
    return S_OK;
}

[[nodiscard]] HRESULT EnsureSchema(sqlite3* db)
{
    int userVersion = 0;
    HRESULT hr      = ReadSingleInt(db, "PRAGMA user_version;", userVersion);
    if (FAILED(hr))
    {
        return hr;
    }

    if (userVersion < 0)
    {
        return E_FAIL;
    }

    if (static_cast<uint32_t>(userVersion) > kSchemaVersion)
    {
        Debug::Error(L"SqliteIndexStore: database schema version {} is newer than supported version {}.", userVersion, kSchemaVersion);
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }

    switch (static_cast<uint32_t>(userVersion))
    {
        case 0u:
        case 1u:
            hr = EnsureSchemaVersion1(db);
            if (FAILED(hr))
            {
                return hr;
            }
            return EnsureSchemaVersion2(db, true);
        case 2u: return EnsureSchemaVersion2(db, false);
        default: return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
}

[[nodiscard]] HRESULT PopulateStoreInfo(sqlite3* db, std::wstring_view normalizedPath, StoreInfo& outInfo)
{
    outInfo.databasePath       = std::wstring(normalizedPath);
    outInfo.databaseBytes      = GetFileBytes(outInfo.databasePath);
    outInfo.writeAheadLogPath  = outInfo.databasePath + L"-wal";
    outInfo.writeAheadLogBytes = GetFileBytes(outInfo.writeAheadLogPath);

    int userVersion = 0;
    HRESULT hr      = ReadSingleInt(db, "PRAGMA user_version;", userVersion);
    if (FAILED(hr))
    {
        return hr;
    }
    outInfo.schemaVersion = static_cast<uint32_t>(std::max(userVersion, 0));

    std::wstring journalMode;
    hr = ReadSingleText(db, "PRAGMA journal_mode;", journalMode);
    if (FAILED(hr))
    {
        return hr;
    }
    outInfo.walEnabled = OrdinalString::EqualsNoCase(journalMode, L"wal");

    int foreignKeys = 0;
    hr              = ReadSingleInt(db, "PRAGMA foreign_keys;", foreignKeys);
    if (FAILED(hr))
    {
        return hr;
    }
    outInfo.foreignKeysEnabled = foreignKeys != 0;

    int autoVacuum = 0;
    hr             = ReadSingleInt(db, "PRAGMA auto_vacuum;", autoVacuum);
    if (FAILED(hr))
    {
        return hr;
    }
    outInfo.incrementalAutoVacuumEnabled = autoVacuum == kSqliteAutoVacuumIncremental;

    hr = ReadPageCounts(db, outInfo.pageCount, outInfo.freelistPageCount);
    if (FAILED(hr))
    {
        return hr;
    }

    bool hasMetaTable            = false;
    bool hasVolumesTable         = false;
    bool hasEntriesTable         = false;
    bool hasNameIndex            = false;
    bool hasExtensionIndex       = false;
    bool hasPathIndex            = false;
    bool hasParentIndex          = false;
    bool hasCreationTimeColumn   = false;
    bool hasLastAccessTimeColumn = false;
    bool hasChangeTimeColumn     = false;
    bool hasAllocationSizeColumn = false;
    hr                           = SchemaObjectExists(db, "table", "meta", hasMetaTable);
    if (SUCCEEDED(hr))
    {
        hr = SchemaObjectExists(db, "table", "volumes", hasVolumesTable);
    }
    if (SUCCEEDED(hr))
    {
        hr = SchemaObjectExists(db, "table", "entries", hasEntriesTable);
    }
    if (SUCCEEDED(hr))
    {
        hr = SchemaObjectExists(db, "index", "idx_entries_name_folded", hasNameIndex);
    }
    if (SUCCEEDED(hr))
    {
        hr = SchemaObjectExists(db, "index", "idx_entries_extension_folded", hasExtensionIndex);
    }
    if (SUCCEEDED(hr))
    {
        hr = SchemaObjectExists(db, "index", "idx_entries_full_path_folded", hasPathIndex);
    }
    if (SUCCEEDED(hr))
    {
        hr = SchemaObjectExists(db, "index", "idx_entries_parent", hasParentIndex);
    }
    if (SUCCEEDED(hr))
    {
        hr = TableColumnExists(db, "entries", "creation_time_100ns", hasCreationTimeColumn);
    }
    if (SUCCEEDED(hr))
    {
        hr = TableColumnExists(db, "entries", "last_access_time_100ns", hasLastAccessTimeColumn);
    }
    if (SUCCEEDED(hr))
    {
        hr = TableColumnExists(db, "entries", "change_time_100ns", hasChangeTimeColumn);
    }
    if (SUCCEEDED(hr))
    {
        hr = TableColumnExists(db, "entries", "allocation_size", hasAllocationSizeColumn);
    }
    if (FAILED(hr))
    {
        return hr;
    }

    bool hasLastCheckpointUtc = false;
    hr                        = ReadMetaText(db, kMetaLastCheckpointUtc, outInfo.lastCheckpointUtc, hasLastCheckpointUtc);
    if (FAILED(hr))
    {
        return hr;
    }
    if (! hasLastCheckpointUtc)
    {
        outInfo.lastCheckpointUtc.clear();
    }

    bool hasLastCompactionUtc = false;
    hr                        = ReadMetaText(db, kMetaLastCompactionUtc, outInfo.lastCompactionUtc, hasLastCompactionUtc);
    if (FAILED(hr))
    {
        return hr;
    }
    if (! hasLastCompactionUtc)
    {
        outInfo.lastCompactionUtc.clear();
    }

    int metaSchemaVersion = 0;
    hr                    = ReadSingleInt(db, "SELECT CAST(value AS INTEGER) FROM meta WHERE key = 'schema_version';", metaSchemaVersion);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReadSingleInt64(db, "SELECT COUNT(*) FROM volumes;", outInfo.volumeCount);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReadSingleInt64(db, "SELECT COUNT(*) FROM entries;", outInfo.entryCount);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReadSingleInt64(db,
                         std::format("SELECT COUNT(*) FROM volumes WHERE state = {};", static_cast<unsigned long long>(kVolumeStateImportedLegacySnapshot)),
                         outInfo.legacyImportVolumeCount);
    if (FAILED(hr))
    {
        return hr;
    }

    outInfo.schemaReady = hasMetaTable && hasVolumesTable && hasEntriesTable && hasNameIndex && hasExtensionIndex && hasPathIndex && hasParentIndex &&
                          hasCreationTimeColumn && hasLastAccessTimeColumn && hasChangeTimeColumn && hasAllocationSizeColumn &&
                          metaSchemaVersion == static_cast<int>(kSchemaVersion) && outInfo.schemaVersion == kSchemaVersion;
    return S_OK;
}

[[nodiscard]] HRESULT FindVolumeId(sqlite3* db, std::wstring_view rootPath, sqlite3_int64& outVolumeId, bool& outExists)
{
    outVolumeId = 0;
    outExists   = false;

    unique_statement statement;
    constexpr std::string_view kSql = "SELECT volume_id FROM volumes WHERE root_path = ?1;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = BindWideText(statement.get(), 1, rootPath);
    if (FAILED(hr))
    {
        Debug::Error(L"SqliteIndexStore: failed to bind root_path for volume lookup '{}'.", std::wstring(rootPath));
        return hr;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult == SQLITE_DONE)
    {
        return S_OK;
    }
    if (sqliteResult != SQLITE_ROW)
    {
        Debug::Error(L"SqliteIndexStore: volume lookup failed for '{}'. code={} message='{}'", std::wstring(rootPath), sqliteResult, GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    outVolumeId = sqlite3_column_int64(statement.get(), 0);
    outExists   = true;
    return S_OK;
}

[[nodiscard]] HRESULT InsertVolume(sqlite3* db, const ReplaceVolumeRequest& request, sqlite3_int64& outVolumeId)
{
    unique_statement statement;
    constexpr std::string_view kSql =
        "INSERT INTO volumes(root_path, fs_kind, journal_id, next_usn, state, entry_count, last_seed_utc, last_replay_utc, last_error_hr) "
        "VALUES(?1, ?2, ?3, ?4, ?5, ?6, NULL, NULL, 0);";
    HRESULT hr = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = BindWideText(statement.get(), 1, request.rootPath);
    if (FAILED(hr))
    {
        Debug::Error(L"SqliteIndexStore: failed to bind root_path for insert '{}'.", request.rootPath);
        return hr;
    }

    const int bindFsKind     = sqlite3_bind_int(statement.get(), 2, static_cast<int>(request.fileSystemKind));
    const int bindJournal    = sqlite3_bind_int64(statement.get(), 3, static_cast<sqlite3_int64>(request.journalId));
    const int bindNextUsn    = sqlite3_bind_int64(statement.get(), 4, static_cast<sqlite3_int64>(request.nextUsn));
    const int bindState      = sqlite3_bind_int(statement.get(), 5, static_cast<int>(request.state));
    const int bindEntryCount = sqlite3_bind_int64(statement.get(), 6, static_cast<sqlite3_int64>(request.entries.size()));
    if (bindFsKind != SQLITE_OK || bindJournal != SQLITE_OK || bindNextUsn != SQLITE_OK || bindState != SQLITE_OK || bindEntryCount != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: failed to bind volume insert for '{}'. code={} message='{}'",
                     request.rootPath,
                     (bindFsKind != SQLITE_OK)    ? bindFsKind
                     : (bindJournal != SQLITE_OK) ? bindJournal
                     : (bindNextUsn != SQLITE_OK) ? bindNextUsn
                     : (bindState != SQLITE_OK)   ? bindState
                                                  : bindEntryCount,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult != SQLITE_DONE)
    {
        Debug::Error(L"SqliteIndexStore: volume insert failed for '{}'. code={} message='{}'", request.rootPath, sqliteResult, GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    outVolumeId = sqlite3_last_insert_rowid(db);
    return S_OK;
}

[[nodiscard]] HRESULT UpdateVolume(sqlite3* db, sqlite3_int64 volumeId, const ReplaceVolumeRequest& request)
{
    unique_statement statement;
    constexpr std::string_view kSql = "UPDATE volumes "
                                      "SET fs_kind = ?1, journal_id = ?2, next_usn = ?3, state = ?4, entry_count = ?5, last_error_hr = 0 "
                                      "WHERE volume_id = ?6;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int bindFsKind     = sqlite3_bind_int(statement.get(), 1, static_cast<int>(request.fileSystemKind));
    const int bindJournal    = sqlite3_bind_int64(statement.get(), 2, static_cast<sqlite3_int64>(request.journalId));
    const int bindNextUsn    = sqlite3_bind_int64(statement.get(), 3, static_cast<sqlite3_int64>(request.nextUsn));
    const int bindState      = sqlite3_bind_int(statement.get(), 4, static_cast<int>(request.state));
    const int bindEntryCount = sqlite3_bind_int64(statement.get(), 5, static_cast<sqlite3_int64>(request.entries.size()));
    const int bindVolumeId   = sqlite3_bind_int64(statement.get(), 6, volumeId);
    if (bindFsKind != SQLITE_OK || bindJournal != SQLITE_OK || bindNextUsn != SQLITE_OK || bindState != SQLITE_OK || bindEntryCount != SQLITE_OK ||
        bindVolumeId != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: failed to bind volume update for '{}'. code={} message='{}'",
                     request.rootPath,
                     (bindFsKind != SQLITE_OK)       ? bindFsKind
                     : (bindJournal != SQLITE_OK)    ? bindJournal
                     : (bindNextUsn != SQLITE_OK)    ? bindNextUsn
                     : (bindState != SQLITE_OK)      ? bindState
                     : (bindEntryCount != SQLITE_OK) ? bindEntryCount
                                                     : bindVolumeId,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult != SQLITE_DONE)
    {
        Debug::Error(L"SqliteIndexStore: volume update failed for '{}'. code={} message='{}'", request.rootPath, sqliteResult, GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    return S_OK;
}

[[nodiscard]] HRESULT DeleteVolumeEntries(sqlite3* db, sqlite3_int64 volumeId)
{
    unique_statement statement;
    constexpr std::string_view kSql = "DELETE FROM entries WHERE volume_id = ?1;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int bindVolumeId = sqlite3_bind_int64(statement.get(), 1, volumeId);
    if (bindVolumeId != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: failed to bind volume_id for entry delete. code={} message='{}'", bindVolumeId, GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult != SQLITE_DONE)
    {
        Debug::Error(L"SqliteIndexStore: entry delete failed for volume_id={}. code={} message='{}'",
                     static_cast<long long>(volumeId),
                     sqliteResult,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    return S_OK;
}

[[nodiscard]] HRESULT DeleteVolumeRow(sqlite3* db, sqlite3_int64 volumeId)
{
    unique_statement statement;
    constexpr std::string_view kSql = "DELETE FROM volumes WHERE volume_id = ?1;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int bindVolumeId = sqlite3_bind_int64(statement.get(), 1, volumeId);
    if (bindVolumeId != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: failed to bind volume_id for volume delete. code={} message='{}'", bindVolumeId, GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult != SQLITE_DONE)
    {
        Debug::Error(L"SqliteIndexStore: volume delete failed for volume_id={}. code={} message='{}'",
                     static_cast<long long>(volumeId),
                     sqliteResult,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    return S_OK;
}

[[nodiscard]] HRESULT InsertVolumeEntries(sqlite3* db, sqlite3_int64 volumeId, const ReplaceVolumeRequest& request)
{
    unique_statement statement;
    constexpr std::string_view kSql =
        "INSERT INTO entries(volume_id, file_id_low, file_id_high, parent_id_low, parent_id_high, full_path, full_path_folded, "
        "name, name_folded, extension_folded, attributes, is_dir, size_bytes, write_time_100ns, creation_time_100ns, last_access_time_100ns, "
        "change_time_100ns, allocation_size) "
        "VALUES(?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18);";
    HRESULT hr = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    for (const ImportedEntry& entry : request.entries)
    {
        static_cast<void>(sqlite3_reset(statement.get()));
        static_cast<void>(sqlite3_clear_bindings(statement.get()));

        const std::wstring fullPathFolded  = FoldText(entry.fullPath);
        const std::wstring nameFolded      = FoldText(entry.name);
        const std::wstring extensionFolded = GetExtensionFolded(entry.name);
        const int bindVolumeId             = sqlite3_bind_int64(statement.get(), 1, volumeId);
        const int bindIdLow                = sqlite3_bind_int64(statement.get(), 2, static_cast<sqlite3_int64>(entry.fileIdLow));
        const int bindIdHigh               = sqlite3_bind_int64(statement.get(), 3, static_cast<sqlite3_int64>(entry.fileIdHigh));
        const int bindParentLow            = sqlite3_bind_int64(statement.get(), 4, static_cast<sqlite3_int64>(entry.parentIdLow));
        const int bindParentHigh           = sqlite3_bind_int64(statement.get(), 5, static_cast<sqlite3_int64>(entry.parentIdHigh));
        const HRESULT bindFullPath         = BindWideText(statement.get(), 6, entry.fullPath);
        const HRESULT bindFullPathFolded   = BindWideText(statement.get(), 7, fullPathFolded);
        const HRESULT bindName             = BindWideText(statement.get(), 8, entry.name);
        const HRESULT bindNameFolded       = BindWideText(statement.get(), 9, nameFolded);
        const HRESULT bindExtensionFolded  = BindWideText(statement.get(), 10, extensionFolded);
        const int bindAttributes           = sqlite3_bind_int64(statement.get(), 11, static_cast<sqlite3_int64>(entry.attributes));
        const int bindIsDir                = sqlite3_bind_int(statement.get(), 12, (entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u ? 1 : 0);
        const int bindSize                 = sqlite3_bind_int64(statement.get(), 13, static_cast<sqlite3_int64>(entry.sizeBytes));
        const int bindWriteTime            = sqlite3_bind_int64(statement.get(), 14, static_cast<sqlite3_int64>(entry.writeTime100ns));
        const int bindCreationTime         = sqlite3_bind_int64(statement.get(), 15, static_cast<sqlite3_int64>(entry.creationTime100ns));
        const int bindLastAccessTime       = sqlite3_bind_int64(statement.get(), 16, static_cast<sqlite3_int64>(entry.lastAccessTime100ns));
        const int bindChangeTime           = sqlite3_bind_int64(statement.get(), 17, static_cast<sqlite3_int64>(entry.changeTime100ns));
        const int bindAllocationSize       = sqlite3_bind_int64(statement.get(), 18, static_cast<sqlite3_int64>(entry.allocationSize));
        if (bindVolumeId != SQLITE_OK || bindIdLow != SQLITE_OK || bindIdHigh != SQLITE_OK || bindParentLow != SQLITE_OK || bindParentHigh != SQLITE_OK ||
            FAILED(bindFullPath) || FAILED(bindFullPathFolded) || FAILED(bindName) || FAILED(bindNameFolded) || FAILED(bindExtensionFolded) ||
            bindAttributes != SQLITE_OK || bindIsDir != SQLITE_OK || bindSize != SQLITE_OK || bindWriteTime != SQLITE_OK || bindCreationTime != SQLITE_OK ||
            bindLastAccessTime != SQLITE_OK || bindChangeTime != SQLITE_OK || bindAllocationSize != SQLITE_OK)
        {
            Debug::Error(L"SqliteIndexStore: failed to bind entry insert for root='{}' entry='{}'. message='{}'",
                         request.rootPath,
                         entry.fullPath,
                         GetSqliteErrorMessage(db));
            return E_FAIL;
        }

        const int sqliteResult = sqlite3_step(statement.get());
        if (sqliteResult != SQLITE_DONE)
        {
            Debug::Error(L"SqliteIndexStore: entry insert failed for root='{}' entry='{}'. code={} message='{}'",
                         request.rootPath,
                         entry.fullPath,
                         sqliteResult,
                         GetSqliteErrorMessage(db));
            return E_FAIL;
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT UpsertVolumeEntries(sqlite3* db, sqlite3_int64 volumeId, std::wstring_view rootPath, const std::vector<ImportedEntry>& entries)
{
    unique_statement statement;
    constexpr std::string_view kSql =
        "INSERT INTO entries(volume_id, file_id_low, file_id_high, parent_id_low, parent_id_high, full_path, full_path_folded, "
        "name, name_folded, extension_folded, attributes, is_dir, size_bytes, write_time_100ns, creation_time_100ns, last_access_time_100ns, "
        "change_time_100ns, allocation_size) "
        "VALUES(?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18) "
        "ON CONFLICT(volume_id, file_id_low, file_id_high) DO UPDATE SET "
        "parent_id_low = excluded.parent_id_low, "
        "parent_id_high = excluded.parent_id_high, "
        "full_path = excluded.full_path, "
        "full_path_folded = excluded.full_path_folded, "
        "name = excluded.name, "
        "name_folded = excluded.name_folded, "
        "extension_folded = excluded.extension_folded, "
        "attributes = excluded.attributes, "
        "is_dir = excluded.is_dir, "
        "size_bytes = excluded.size_bytes, "
        "write_time_100ns = excluded.write_time_100ns, "
        "creation_time_100ns = excluded.creation_time_100ns, "
        "last_access_time_100ns = excluded.last_access_time_100ns, "
        "change_time_100ns = excluded.change_time_100ns, "
        "allocation_size = excluded.allocation_size;";
    HRESULT hr = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    for (const ImportedEntry& entry : entries)
    {
        static_cast<void>(sqlite3_reset(statement.get()));
        static_cast<void>(sqlite3_clear_bindings(statement.get()));

        const std::wstring fullPathFolded  = FoldText(entry.fullPath);
        const std::wstring nameFolded      = FoldText(entry.name);
        const std::wstring extensionFolded = GetExtensionFolded(entry.name);
        const int bindVolumeId             = sqlite3_bind_int64(statement.get(), 1, volumeId);
        const int bindIdLow                = sqlite3_bind_int64(statement.get(), 2, static_cast<sqlite3_int64>(entry.fileIdLow));
        const int bindIdHigh               = sqlite3_bind_int64(statement.get(), 3, static_cast<sqlite3_int64>(entry.fileIdHigh));
        const int bindParentLow            = sqlite3_bind_int64(statement.get(), 4, static_cast<sqlite3_int64>(entry.parentIdLow));
        const int bindParentHigh           = sqlite3_bind_int64(statement.get(), 5, static_cast<sqlite3_int64>(entry.parentIdHigh));
        const HRESULT bindFullPath         = BindWideText(statement.get(), 6, entry.fullPath);
        const HRESULT bindFullPathFolded   = BindWideText(statement.get(), 7, fullPathFolded);
        const HRESULT bindName             = BindWideText(statement.get(), 8, entry.name);
        const HRESULT bindNameFolded       = BindWideText(statement.get(), 9, nameFolded);
        const HRESULT bindExtensionFolded  = BindWideText(statement.get(), 10, extensionFolded);
        const int bindAttributes           = sqlite3_bind_int64(statement.get(), 11, static_cast<sqlite3_int64>(entry.attributes));
        const int bindIsDir                = sqlite3_bind_int(statement.get(), 12, (entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u ? 1 : 0);
        const int bindSize                 = sqlite3_bind_int64(statement.get(), 13, static_cast<sqlite3_int64>(entry.sizeBytes));
        const int bindWriteTime            = sqlite3_bind_int64(statement.get(), 14, static_cast<sqlite3_int64>(entry.writeTime100ns));
        const int bindCreationTime         = sqlite3_bind_int64(statement.get(), 15, static_cast<sqlite3_int64>(entry.creationTime100ns));
        const int bindLastAccessTime       = sqlite3_bind_int64(statement.get(), 16, static_cast<sqlite3_int64>(entry.lastAccessTime100ns));
        const int bindChangeTime           = sqlite3_bind_int64(statement.get(), 17, static_cast<sqlite3_int64>(entry.changeTime100ns));
        const int bindAllocationSize       = sqlite3_bind_int64(statement.get(), 18, static_cast<sqlite3_int64>(entry.allocationSize));
        if (bindVolumeId != SQLITE_OK || bindIdLow != SQLITE_OK || bindIdHigh != SQLITE_OK || bindParentLow != SQLITE_OK || bindParentHigh != SQLITE_OK ||
            FAILED(bindFullPath) || FAILED(bindFullPathFolded) || FAILED(bindName) || FAILED(bindNameFolded) || FAILED(bindExtensionFolded) ||
            bindAttributes != SQLITE_OK || bindIsDir != SQLITE_OK || bindSize != SQLITE_OK || bindWriteTime != SQLITE_OK || bindCreationTime != SQLITE_OK ||
            bindLastAccessTime != SQLITE_OK || bindChangeTime != SQLITE_OK || bindAllocationSize != SQLITE_OK)
        {
            Debug::Error(L"SqliteIndexStore: failed to bind entry upsert for root='{}' entry='{}'. message='{}'",
                         std::wstring(rootPath),
                         entry.fullPath,
                         GetSqliteErrorMessage(db));
            return E_FAIL;
        }

        const int sqliteResult = sqlite3_step(statement.get());
        if (sqliteResult != SQLITE_DONE)
        {
            Debug::Error(L"SqliteIndexStore: entry upsert failed for root='{}' entry='{}'. code={} message='{}'",
                         std::wstring(rootPath),
                         entry.fullPath,
                         sqliteResult,
                         GetSqliteErrorMessage(db));
            return E_FAIL;
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT DeleteVolumeEntriesById(sqlite3* db, sqlite3_int64 volumeId, std::span<const DeletedEntryId> entryIds)
{
    if (entryIds.empty())
    {
        return S_OK;
    }

    unique_statement statement;
    constexpr std::string_view kSql = "DELETE FROM entries WHERE volume_id = ?1 AND file_id_low = ?2 AND file_id_high = ?3;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    for (const DeletedEntryId& entryId : entryIds)
    {
        sqlite3_reset(statement.get());
        sqlite3_clear_bindings(statement.get());

        const int bindVolumeId = sqlite3_bind_int64(statement.get(), 1, volumeId);
        const int bindIdLow    = sqlite3_bind_int64(statement.get(), 2, static_cast<sqlite3_int64>(entryId.fileIdLow));
        const int bindIdHigh   = sqlite3_bind_int64(statement.get(), 3, static_cast<sqlite3_int64>(entryId.fileIdHigh));
        if (bindVolumeId != SQLITE_OK || bindIdLow != SQLITE_OK || bindIdHigh != SQLITE_OK)
        {
            Debug::Error(L"SqliteIndexStore: failed to bind entry delete for volume_id={}. code={} message='{}'",
                         static_cast<long long>(volumeId),
                         (bindVolumeId != SQLITE_OK) ? bindVolumeId
                         : (bindIdLow != SQLITE_OK)  ? bindIdLow
                                                     : bindIdHigh,
                         GetSqliteErrorMessage(db));
            return E_FAIL;
        }

        const int sqliteResult = sqlite3_step(statement.get());
        if (sqliteResult != SQLITE_DONE)
        {
            Debug::Error(L"SqliteIndexStore: entry delete failed for volume_id={}. code={} message='{}'",
                         static_cast<long long>(volumeId),
                         sqliteResult,
                         GetSqliteErrorMessage(db));
            return E_FAIL;
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT ReadVolumeEntryCount(sqlite3* db, sqlite3_int64 volumeId, uint64_t& outCount)
{
    outCount = 0u;

    unique_statement statement;
    constexpr std::string_view kSql = "SELECT COUNT(*) FROM entries WHERE volume_id = ?1;";
    HRESULT hr                      = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int bindVolumeId = sqlite3_bind_int64(statement.get(), 1, volumeId);
    if (bindVolumeId != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: failed to bind volume_id for entry count. code={} message='{}'", bindVolumeId, GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    const int sqliteResult = sqlite3_step(statement.get());
    if (sqliteResult != SQLITE_ROW)
    {
        Debug::Error(L"SqliteIndexStore: entry count query failed for volume_id={}. code={} message='{}'",
                     static_cast<long long>(volumeId),
                     sqliteResult,
                     GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    outCount = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 0));
    return S_OK;
}

[[nodiscard]] HRESULT ReadVolumeEntries(sqlite3* db, sqlite3_int64 volumeId, std::vector<ImportedEntry>& outEntries)
{
    outEntries.clear();

    unique_statement statement;
    constexpr std::string_view kSql =
        "SELECT file_id_low, file_id_high, parent_id_low, parent_id_high, full_path, name, attributes, size_bytes, write_time_100ns, "
        "creation_time_100ns, last_access_time_100ns, change_time_100ns, allocation_size "
        "FROM entries WHERE volume_id = ?1 ORDER BY full_path_folded;";
    HRESULT hr = PrepareStatement(db, kSql, statement);
    if (FAILED(hr))
    {
        return hr;
    }

    const int bindVolumeId = sqlite3_bind_int64(statement.get(), 1, volumeId);
    if (bindVolumeId != SQLITE_OK)
    {
        Debug::Error(L"SqliteIndexStore: failed to bind volume_id for entry load. code={} message='{}'", bindVolumeId, GetSqliteErrorMessage(db));
        return E_FAIL;
    }

    while (true)
    {
        const int sqliteResult = sqlite3_step(statement.get());
        if (sqliteResult == SQLITE_DONE)
        {
            return S_OK;
        }
        if (sqliteResult != SQLITE_ROW)
        {
            Debug::Error(L"SqliteIndexStore: entry load failed for volume_id={}. code={} message='{}'",
                         static_cast<long long>(volumeId),
                         sqliteResult,
                         GetSqliteErrorMessage(db));
            return E_FAIL;
        }

        outEntries.push_back({
            .fileIdLow           = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 0)),
            .fileIdHigh          = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 1)),
            .parentIdLow         = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 2)),
            .parentIdHigh        = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 3)),
            .fullPath            = ReadWideColumn(statement.get(), 4),
            .name                = ReadWideColumn(statement.get(), 5),
            .attributes          = static_cast<unsigned long>(sqlite3_column_int64(statement.get(), 6)),
            .sizeBytes           = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 7)),
            .writeTime100ns      = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 8)),
            .creationTime100ns   = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 9)),
            .lastAccessTime100ns = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 10)),
            .changeTime100ns     = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 11)),
            .allocationSize      = static_cast<uint64_t>(sqlite3_column_int64(statement.get(), 12)),
        });
    }
}
} // namespace

HRESULT EnsureBootstrap(std::wstring_view databasePath, StoreInfo* outInfo) noexcept
{
    try
    {
        const std::wstring normalizedPath = NormalizeDatabasePath(databasePath);
        if (normalizedPath.empty())
        {
            return E_INVALIDARG;
        }

        const HRESULT directoryHr = EnsureParentDirectory(normalizedPath);
        if (FAILED(directoryHr))
        {
            return directoryHr;
        }

        unique_connection db;
        HRESULT hr = OpenConnection(normalizedPath, SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE | SQLITE_OPEN_NOMUTEX, db);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ApplyPragmas(db.get());
        if (FAILED(hr))
        {
            return hr;
        }

        hr = EnsureSchema(db.get());
        if (FAILED(hr))
        {
            return hr;
        }

        if (outInfo == nullptr)
        {
            return S_OK;
        }

        return PopulateStoreInfo(db.get(), normalizedPath, *outInfo);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SqliteIndexStore::EnsureBootstrap: std::exception");
        return E_FAIL;
    }
}

HRESULT InspectStore(std::wstring_view databasePath, StoreInfo& outInfo) noexcept
{
    try
    {
        outInfo = {};

        const std::wstring normalizedPath = NormalizeDatabasePath(databasePath);
        if (normalizedPath.empty())
        {
            return E_INVALIDARG;
        }

        std::error_code existsEc;
        const bool exists = std::filesystem::exists(std::filesystem::path(normalizedPath), existsEc);
        if (existsEc)
        {
            return HRESULT_FROM_WIN32(static_cast<unsigned long>(existsEc.value()));
        }
        if (! exists)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        unique_connection db;
        const HRESULT hr = OpenConnection(normalizedPath, SQLITE_OPEN_READONLY | SQLITE_OPEN_NOMUTEX, db);
        if (FAILED(hr))
        {
            return hr;
        }

        return PopulateStoreInfo(db.get(), normalizedPath, outInfo);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SqliteIndexStore::InspectStore: std::exception");
        return E_FAIL;
    }
}

HRESULT InspectVolume(std::wstring_view databasePath, std::wstring_view rootPath, VolumeInfo& outInfo) noexcept
{
    try
    {
        outInfo = {};
        if (rootPath.empty())
        {
            return E_INVALIDARG;
        }

        const HRESULT bootstrapHr = EnsureBootstrap(databasePath, nullptr);
        if (FAILED(bootstrapHr))
        {
            return bootstrapHr;
        }

        const std::wstring normalizedPath = NormalizeDatabasePath(databasePath);
        unique_connection db;
        HRESULT hr = OpenConnection(normalizedPath, SQLITE_OPEN_READONLY | SQLITE_OPEN_NOMUTEX, db);
        if (FAILED(hr))
        {
            return hr;
        }

        VolumeRow volumeRow{};
        hr = ReadVolumeRow(db.get(), rootPath, volumeRow);
        if (FAILED(hr))
        {
            return hr;
        }

        outInfo.fileSystemKind = volumeRow.fileSystemKind;
        outInfo.journalId      = volumeRow.journalId;
        outInfo.nextUsn        = volumeRow.nextUsn;
        outInfo.entryCount     = volumeRow.entryCount;
        outInfo.state          = volumeRow.state;
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SqliteIndexStore: InspectVolume failed with an unexpected std::exception.");
        outInfo = {};
        return E_FAIL;
    }
}

HRESULT LoadVolume(std::wstring_view databasePath, std::wstring_view rootPath, ReplaceVolumeRequest& outVolume) noexcept
{
    try
    {
        outVolume = {};
        if (rootPath.empty())
        {
            return E_INVALIDARG;
        }

        const HRESULT bootstrapHr = EnsureBootstrap(databasePath, nullptr);
        if (FAILED(bootstrapHr))
        {
            return bootstrapHr;
        }

        const std::wstring normalizedPath = NormalizeDatabasePath(databasePath);
        unique_connection db;
        HRESULT hr = OpenConnection(normalizedPath, SQLITE_OPEN_READONLY | SQLITE_OPEN_NOMUTEX, db);
        if (FAILED(hr))
        {
            return hr;
        }

        VolumeRow volumeRow{};
        hr = ReadVolumeRow(db.get(), rootPath, volumeRow);
        if (FAILED(hr))
        {
            return hr;
        }

        outVolume.rootPath       = std::wstring(rootPath);
        outVolume.fileSystemKind = volumeRow.fileSystemKind;
        outVolume.journalId      = volumeRow.journalId;
        outVolume.nextUsn        = volumeRow.nextUsn;
        outVolume.state          = volumeRow.state;
        return ReadVolumeEntries(db.get(), volumeRow.volumeId, outVolume.entries);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SqliteIndexStore::LoadVolume: std::exception");
        outVolume = {};
        return E_FAIL;
    }
}

HRESULT ReplaceVolume(std::wstring_view databasePath, const ReplaceVolumeRequest& request, ReplaceVolumeResult* outResult) noexcept
{
    try
    {
        if (request.rootPath.empty())
        {
            return E_INVALIDARG;
        }

        const HRESULT bootstrapHr = EnsureBootstrap(databasePath, nullptr);
        if (FAILED(bootstrapHr))
        {
            return bootstrapHr;
        }

        const std::wstring normalizedPath = NormalizeDatabasePath(databasePath);
        unique_connection db;
        HRESULT hr = OpenConnection(normalizedPath, SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE | SQLITE_OPEN_NOMUTEX, db);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ExecuteSql(db.get(), "BEGIN IMMEDIATE;", L"BEGIN IMMEDIATE");
        if (FAILED(hr))
        {
            return hr;
        }

        bool commit         = false;
        const auto rollback = wil::scope_exit([&]() noexcept
        {
            if (! commit)
            {
                static_cast<void>(ExecuteSql(db.get(), "ROLLBACK;", L"ROLLBACK"));
            }
        });

        sqlite3_int64 volumeId = 0;
        bool exists            = false;
        hr                     = FindVolumeId(db.get(), request.rootPath, volumeId, exists);
        if (FAILED(hr))
        {
            return hr;
        }

        ReplaceVolumeResult result{};
        if (! exists)
        {
            hr = InsertVolume(db.get(), request, volumeId);
            if (FAILED(hr))
            {
                return hr;
            }
            result.insertedNewVolume = true;
        }
        else
        {
            hr = UpdateVolume(db.get(), volumeId, request);
            if (FAILED(hr))
            {
                return hr;
            }

            hr = DeleteVolumeEntries(db.get(), volumeId);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        hr = InsertVolumeEntries(db.get(), volumeId, request);
        if (FAILED(hr))
        {
            return hr;
        }

        if (! result.insertedNewVolume)
        {
            hr = UpdateVolume(db.get(), volumeId, request);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        hr = ExecuteSql(db.get(), "COMMIT;", L"COMMIT");
        if (FAILED(hr))
        {
            return hr;
        }

        commit                    = true;
        result.importedEntryCount = static_cast<uint64_t>(request.entries.size());
        if (outResult != nullptr)
        {
            *outResult = result;
        }

        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SqliteIndexStore::ReplaceVolume: std::exception");
        return E_FAIL;
    }
}

HRESULT ApplyJournalDelta(std::wstring_view databasePath, const ApplyJournalDeltaRequest& request, ApplyJournalDeltaResult* outResult) noexcept
{
    try
    {
        if (request.rootPath.empty())
        {
            return E_INVALIDARG;
        }

        const HRESULT bootstrapHr = EnsureBootstrap(databasePath, nullptr);
        if (FAILED(bootstrapHr))
        {
            return bootstrapHr;
        }

        const std::wstring normalizedPath = NormalizeDatabasePath(databasePath);
        unique_connection db;
        HRESULT hr = OpenConnection(normalizedPath, SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE | SQLITE_OPEN_NOMUTEX, db);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ExecuteSql(db.get(), "BEGIN IMMEDIATE;", L"BEGIN IMMEDIATE");
        if (FAILED(hr))
        {
            return hr;
        }

        bool commit         = false;
        const auto rollback = wil::scope_exit([&]() noexcept
        {
            if (! commit)
            {
                static_cast<void>(ExecuteSql(db.get(), "ROLLBACK;", L"ROLLBACK"));
            }
        });

        sqlite3_int64 volumeId = 0;
        bool exists            = false;
        hr                     = FindVolumeId(db.get(), request.rootPath, volumeId, exists);
        if (FAILED(hr))
        {
            return hr;
        }

        ApplyJournalDeltaResult result{};
        bool seededMissingVolumeFromSnapshot = false;
        if (! exists)
        {
            ReplaceVolumeRequest seed{};
            seed.rootPath       = request.rootPath;
            seed.fileSystemKind = request.fileSystemKind;
            seed.journalId      = request.journalId;
            seed.nextUsn        = request.nextUsn;
            seed.state          = request.seedStateIfMissing;
            hr                  = InsertVolume(db.get(), seed, volumeId);
            if (FAILED(hr))
            {
                return hr;
            }
            result.insertedNewVolume = true;

            if (! request.seedEntriesIfMissing.empty())
            {
                hr = UpsertVolumeEntries(db.get(), volumeId, request.rootPath, request.seedEntriesIfMissing);
                if (FAILED(hr))
                {
                    return hr;
                }
                seededMissingVolumeFromSnapshot = true;
            }
        }

        if (! seededMissingVolumeFromSnapshot)
        {
            hr = DeleteVolumeEntriesById(db.get(), volumeId, request.deletedEntries);
            if (FAILED(hr))
            {
                return hr;
            }

            hr = UpsertVolumeEntries(db.get(), volumeId, request.rootPath, request.upsertEntries);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        uint64_t entryCount = 0u;
        hr                  = ReadVolumeEntryCount(db.get(), volumeId, entryCount);
        if (FAILED(hr))
        {
            return hr;
        }

        unique_statement statement;
        constexpr std::string_view kUpdateSql =
            "UPDATE volumes "
            "SET fs_kind = ?1, journal_id = ?2, next_usn = ?3, state = ?4, entry_count = ?5, last_replay_utc = ?6, last_error_hr = 0 "
            "WHERE volume_id = ?7;";
        hr = PrepareStatement(db.get(), kUpdateSql, statement);
        if (FAILED(hr))
        {
            return hr;
        }

        const std::wstring replayUtc = GetCurrentUtcTimestampText();
        const int bindFsKind         = sqlite3_bind_int(statement.get(), 1, static_cast<int>(request.fileSystemKind));
        const int bindJournal        = sqlite3_bind_int64(statement.get(), 2, static_cast<sqlite3_int64>(request.journalId));
        const int bindNextUsn        = sqlite3_bind_int64(statement.get(), 3, static_cast<sqlite3_int64>(request.nextUsn));
        const int bindState          = sqlite3_bind_int(statement.get(), 4, static_cast<int>(request.state));
        const int bindEntryCount     = sqlite3_bind_int64(statement.get(), 5, static_cast<sqlite3_int64>(entryCount));
        const HRESULT bindReplayUtc  = BindWideText(statement.get(), 6, replayUtc);
        const int bindVolumeId       = sqlite3_bind_int64(statement.get(), 7, volumeId);
        if (bindFsKind != SQLITE_OK || bindJournal != SQLITE_OK || bindNextUsn != SQLITE_OK || bindState != SQLITE_OK || bindEntryCount != SQLITE_OK ||
            FAILED(bindReplayUtc) || bindVolumeId != SQLITE_OK)
        {
            Debug::Error(
                L"SqliteIndexStore: failed to bind journal delta volume update for '{}'. message='{}'", request.rootPath, GetSqliteErrorMessage(db.get()));
            return E_FAIL;
        }

        const int sqliteResult = sqlite3_step(statement.get());
        if (sqliteResult != SQLITE_DONE)
        {
            Debug::Error(L"SqliteIndexStore: journal delta volume update failed for '{}'. code={} message='{}'",
                         request.rootPath,
                         sqliteResult,
                         GetSqliteErrorMessage(db.get()));
            return E_FAIL;
        }

        hr = ExecuteSql(db.get(), "COMMIT;", L"COMMIT");
        if (FAILED(hr))
        {
            return hr;
        }

        commit                    = true;
        result.deletedEntryCount  = static_cast<uint64_t>(request.deletedEntries.size());
        result.upsertedEntryCount = static_cast<uint64_t>(request.upsertEntries.size());
        if (outResult != nullptr)
        {
            *outResult = result;
        }

        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SqliteIndexStore::ApplyJournalDelta: std::exception");
        return E_FAIL;
    }
}

HRESULT DeleteVolume(std::wstring_view databasePath, std::wstring_view rootPath) noexcept
{
    try
    {
        if (rootPath.empty())
        {
            return E_INVALIDARG;
        }

        const HRESULT bootstrapHr = EnsureBootstrap(databasePath, nullptr);
        if (FAILED(bootstrapHr))
        {
            return bootstrapHr;
        }

        const std::wstring normalizedPath = NormalizeDatabasePath(databasePath);
        unique_connection db;
        HRESULT hr = OpenConnection(normalizedPath, SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE | SQLITE_OPEN_NOMUTEX, db);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ExecuteSql(db.get(), "BEGIN IMMEDIATE;", L"BEGIN IMMEDIATE");
        if (FAILED(hr))
        {
            return hr;
        }

        bool commit         = false;
        const auto rollback = wil::scope_exit([&]() noexcept
        {
            if (! commit)
            {
                static_cast<void>(ExecuteSql(db.get(), "ROLLBACK;", L"ROLLBACK"));
            }
        });

        sqlite3_int64 volumeId = 0;
        bool exists            = false;
        hr                     = FindVolumeId(db.get(), rootPath, volumeId, exists);
        if (FAILED(hr))
        {
            return hr;
        }

        if (exists)
        {
            hr = DeleteVolumeEntries(db.get(), volumeId);
            if (FAILED(hr))
            {
                return hr;
            }

            hr = DeleteVolumeRow(db.get(), volumeId);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        hr = ExecuteSql(db.get(), "COMMIT;", L"COMMIT");
        if (FAILED(hr))
        {
            return hr;
        }

        commit = true;
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SqliteIndexStore::DeleteVolume: std::exception");
        return E_FAIL;
    }
}

HRESULT RunManualMaintenance(std::wstring_view databasePath, ManualMaintenanceResult* outResult) noexcept
{
    try
    {
        const HRESULT bootstrapHr = EnsureBootstrap(databasePath, nullptr);
        if (FAILED(bootstrapHr))
        {
            return bootstrapHr;
        }

        const std::wstring normalizedPath = NormalizeDatabasePath(databasePath);
        unique_connection db;
        HRESULT hr = OpenConnection(normalizedPath, SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE | SQLITE_OPEN_NOMUTEX, db);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ApplyPragmas(db.get());
        if (FAILED(hr))
        {
            return hr;
        }

        hr = EnsureSchema(db.get());
        if (FAILED(hr))
        {
            return hr;
        }

        ManualMaintenanceResult result{};
        hr = PopulateStoreInfo(db.get(), normalizedPath, result.before);
        if (FAILED(hr))
        {
            return hr;
        }

        const std::wstring checkpointUtcBeforeVacuum = GetCurrentUtcTimestampText();
        hr                                           = RunWalCheckpoint(db.get(), SQLITE_CHECKPOINT_TRUNCATE, L"manual WAL checkpoint");
        if (FAILED(hr))
        {
            return hr;
        }

        hr = UpsertMetaText(db.get(), kMetaLastCheckpointUtc, checkpointUtcBeforeVacuum);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ExecuteSql(db.get(), "VACUUM;", L"VACUUM");
        if (FAILED(hr))
        {
            return hr;
        }
        result.ranVacuum = true;

        const std::wstring compactionUtc = GetCurrentUtcTimestampText();
        hr                               = UpsertMetaText(db.get(), kMetaLastCompactionUtc, compactionUtc);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = RunWalCheckpoint(db.get(), SQLITE_CHECKPOINT_TRUNCATE, L"post-VACUUM WAL checkpoint");
        if (FAILED(hr))
        {
            return hr;
        }

        hr = UpsertMetaText(db.get(), kMetaLastCheckpointUtc, compactionUtc);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = PopulateStoreInfo(db.get(), normalizedPath, result.after);
        if (FAILED(hr))
        {
            return hr;
        }

        if (outResult != nullptr)
        {
            *outResult = std::move(result);
        }
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SqliteIndexStore::RunManualMaintenance: std::exception");
        return E_FAIL;
    }
}

HRESULT RunAutomaticMaintenance(std::wstring_view databasePath,
                                const LocalSearchIndexCore::SqliteMaintenancePolicy& policy,
                                AutomaticMaintenanceResult* outResult) noexcept
{
    try
    {
        if (outResult != nullptr)
        {
            *outResult = {};
        }

        const HRESULT bootstrapHr = EnsureBootstrap(databasePath, nullptr);
        if (FAILED(bootstrapHr))
        {
            return bootstrapHr;
        }

        const std::wstring normalizedPath = NormalizeDatabasePath(databasePath);
        unique_connection db;
        HRESULT hr = OpenConnection(normalizedPath, SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE | SQLITE_OPEN_NOMUTEX, db);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ApplyPragmas(db.get());
        if (FAILED(hr))
        {
            return hr;
        }

        hr = EnsureSchema(db.get());
        if (FAILED(hr))
        {
            return hr;
        }

        AutomaticMaintenanceResult result{};
        hr = PopulateStoreInfo(db.get(), normalizedPath, result.before);
        if (FAILED(hr))
        {
            return hr;
        }

        uint64_t pageSizeBytes = 0u;
        hr                     = ReadPageSize(db.get(), pageSizeBytes);
        if (FAILED(hr))
        {
            return hr;
        }

        const bool shouldCheckpoint         = policy.autoCheckpointTargetBytes != 0u && result.before.writeAheadLogBytes >= policy.autoCheckpointTargetBytes;
        const uint64_t fragmentationPercent = CalculateFragmentationPercent(result.before);
        const uint64_t reclaimableBytes     = CalculateReclaimableBytes(result.before, pageSizeBytes);
        const bool shouldCompact = result.before.freelistPageCount != 0u &&
                                   (fragmentationPercent >= policy.autoCompactionFragmentationPercent || reclaimableBytes >= policy.autoCompactionMinBytes);

        result.maintenanceNeeded = shouldCheckpoint || shouldCompact;
        if (! result.maintenanceNeeded)
        {
            result.after = result.before;
            if (outResult != nullptr)
            {
                *outResult = std::move(result);
            }
            return S_FALSE;
        }

        if (shouldCheckpoint || shouldCompact)
        {
            const std::wstring checkpointUtc = GetCurrentUtcTimestampText();
            hr                               = RunWalCheckpoint(db.get(), SQLITE_CHECKPOINT_TRUNCATE, L"automatic WAL checkpoint");
            if (FAILED(hr))
            {
                return hr;
            }

            hr = UpsertMetaText(db.get(), kMetaLastCheckpointUtc, checkpointUtc);
            if (FAILED(hr))
            {
                return hr;
            }
            result.ranCheckpoint = true;
        }

        if (shouldCompact)
        {
            result.requestedVacuumPages = std::min<uint64_t>(result.before.freelistPageCount, kAutomaticIncrementalVacuumMaxPages);
            if (result.requestedVacuumPages != 0u)
            {
                hr = ExecuteSql(db.get(), std::format("PRAGMA incremental_vacuum({});", result.requestedVacuumPages), L"automatic incremental vacuum");
                if (FAILED(hr))
                {
                    return hr;
                }
                result.ranIncrementalVacuum = true;
            }

            const std::wstring compactionUtc = GetCurrentUtcTimestampText();
            hr                               = UpsertMetaText(db.get(), kMetaLastCompactionUtc, compactionUtc);
            if (FAILED(hr))
            {
                return hr;
            }

            hr = RunWalCheckpoint(db.get(), SQLITE_CHECKPOINT_TRUNCATE, L"post-incremental-vacuum WAL checkpoint");
            if (FAILED(hr))
            {
                return hr;
            }

            hr = UpsertMetaText(db.get(), kMetaLastCheckpointUtc, compactionUtc);
            if (FAILED(hr))
            {
                return hr;
            }
            result.ranCheckpoint = true;
        }

        if (result.ranCheckpoint)
        {
            // Metadata updates above dirty the WAL again; finish maintenance with a final truncate checkpoint
            // so the archived post-maintenance store shape reflects the actual checkpointed state.
            hr = RunWalCheckpoint(db.get(), SQLITE_CHECKPOINT_TRUNCATE, L"final automatic WAL checkpoint");
            if (FAILED(hr))
            {
                return hr;
            }
        }

        hr = PopulateStoreInfo(db.get(), normalizedPath, result.after);
        if (FAILED(hr))
        {
            return hr;
        }

        if (result.before.freelistPageCount > result.after.freelistPageCount)
        {
            result.reclaimedVacuumPages = result.before.freelistPageCount - result.after.freelistPageCount;
        }

        if (outResult != nullptr)
        {
            *outResult = std::move(result);
        }
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SqliteIndexStore::RunAutomaticMaintenance: std::exception");
        return E_FAIL;
    }
}

HRESULT EnumerateVolume(std::wstring_view databasePath,
                        const QueryRequest& request,
                        LocalSearchIndexCore::CancelCheckFn cancelCheck,
                        void* cancelCookie,
                        LocalSearchIndexCore::CandidateCallbackFn candidateCallback,
                        void* candidateCookie,
                        QueryRuntimeStats* outStats) noexcept
{
    try
    {
        if (outStats != nullptr)
        {
            *outStats = {};
        }

        if (request.rootPath.empty() || candidateCallback == nullptr)
        {
            return E_INVALIDARG;
        }

        if (! request.includeFiles && ! request.includeDirectories)
        {
            return S_OK;
        }

        const HRESULT bootstrapHr = EnsureBootstrap(databasePath, nullptr);
        if (FAILED(bootstrapHr))
        {
            return bootstrapHr;
        }

        const std::wstring normalizedPath = NormalizeDatabasePath(databasePath);
        unique_connection db;
        HRESULT hr = OpenConnection(normalizedPath, SQLITE_OPEN_READONLY | SQLITE_OPEN_NOMUTEX, db);
        if (FAILED(hr))
        {
            return hr;
        }

        VolumeRow volumeRow{};
        hr = ReadVolumeRow(db.get(), request.rootPath, volumeRow);
        if (FAILED(hr))
        {
            return hr;
        }

        RootEntryRow rootEntry{};
        hr = ReadRootEntry(db.get(), volumeRow.volumeId, request.rootPath, rootEntry);
        if (FAILED(hr))
        {
            return hr;
        }

        QueryRuntimeStats runtimeStats{};
        runtimeStats.volumeReady        = volumeRow.state == kVolumeStateReady;
        runtimeStats.readOnlyConnection = true;
        runtimeStats.fileSystemKind     = volumeRow.fileSystemKind;
        runtimeStats.entryCount         = volumeRow.entryCount;
        runtimeStats.journalId          = volumeRow.journalId;
        runtimeStats.nextUsn            = volumeRow.nextUsn;
#ifdef ENABLE_TESTS
        const uint64_t injectedFailAfterEmittedRows = GetInjectedEnumerateFailureAfterEmittedRows();
#endif

        hr = ReadVolumeTypeCounts(db.get(), volumeRow.volumeId, runtimeStats.fileCount, runtimeStats.directoryCount);
        if (FAILED(hr))
        {
            return hr;
        }

        if ((rootEntry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
        {
            runtimeStats.entryCount     = (runtimeStats.entryCount > 0u) ? (runtimeStats.entryCount - 1u) : 0u;
            runtimeStats.directoryCount = (runtimeStats.directoryCount > 0u) ? (runtimeStats.directoryCount - 1u) : 0u;
        }
        else if (runtimeStats.fileCount > 0u)
        {
            runtimeStats.fileCount -= 1u;
        }

        NamePrefilter prefilter{};
        runtimeStats.usedNamePrefilter = TryBuildNamePrefilter(request, prefilter);

        unique_statement statement;
        hr = PrepareEnumerateStatement(db.get(), volumeRow.volumeId, rootEntry, request, prefilter, statement);
        if (FAILED(hr))
        {
            return hr;
        }

        while (true)
        {
            if ((runtimeStats.scannedRows % kQueryCancelCheckInterval) == 0u)
            {
                hr = CheckCancelled(cancelCheck, cancelCookie);
                if (FAILED(hr))
                {
                    return hr;
                }
            }

            const int sqliteResult = sqlite3_step(statement.get());
            if (sqliteResult == SQLITE_DONE)
            {
                break;
            }
            if (sqliteResult != SQLITE_ROW)
            {
                Debug::Error(L"SqliteIndexStore: query iteration failed for '{}'. code={} message='{}'",
                             request.rootPath,
                             sqliteResult,
                             GetSqliteErrorMessage(db.get()));
                return E_FAIL;
            }

            ++runtimeStats.scannedRows;

            LocalSearchIndexCore::Candidate candidate{};
            candidate.fullPath            = ReadWideColumn(statement.get(), 0);
            const std::wstring storedName = ReadWideColumn(statement.get(), 1);
            candidate.displayName         = GetDisplayName(candidate.fullPath, storedName);
            candidate.fileAttributes      = static_cast<unsigned long>(sqlite3_column_int64(statement.get(), 2));
            if (runtimeStats.volumeReady)
            {
                candidate.metadataFlags = LocalSearchIndexCore::CANDIDATE_METADATA_END_OF_FILE | LocalSearchIndexCore::CANDIDATE_METADATA_LAST_WRITE_TIME |
                                          LocalSearchIndexCore::CANDIDATE_METADATA_CREATION_TIME | LocalSearchIndexCore::CANDIDATE_METADATA_LAST_ACCESS_TIME |
                                          LocalSearchIndexCore::CANDIDATE_METADATA_CHANGE_TIME | LocalSearchIndexCore::CANDIDATE_METADATA_ALLOCATION_SIZE;
                candidate.endOfFile     = sqlite3_column_int64(statement.get(), 3);
                candidate.lastWriteTime100ns  = sqlite3_column_int64(statement.get(), 4);
                candidate.creationTime100ns   = sqlite3_column_int64(statement.get(), 5);
                candidate.lastAccessTime100ns = sqlite3_column_int64(statement.get(), 6);
                candidate.changeTime100ns     = sqlite3_column_int64(statement.get(), 7);
                candidate.allocationSize      = sqlite3_column_int64(statement.get(), 8);
            }

            hr = candidateCallback(&candidate, candidateCookie);
            if (hr == S_FALSE)
            {
                break;
            }
            if (FAILED(hr))
            {
                return hr;
            }

            ++runtimeStats.emittedRows;
#ifdef ENABLE_TESTS
            if (injectedFailAfterEmittedRows != 0u && runtimeStats.emittedRows >= injectedFailAfterEmittedRows)
            {
                Debug::Warning(
                    L"SqliteIndexStore: injected enumerate failure after {} emitted row(s) for '{}'.", injectedFailAfterEmittedRows, request.rootPath);
                return E_FAIL;
            }
#endif
        }

        if (outStats != nullptr)
        {
            *outStats = runtimeStats;
        }

        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SqliteIndexStore::EnumerateVolume: std::exception");
        return E_FAIL;
    }
}
} // namespace SqliteIndexStore
