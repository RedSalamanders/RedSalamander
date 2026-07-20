#include "ViewerSqlite.Engine.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <chrono>
#include <cstring>
#include <cwctype>
#include <format>
#include <limits>
#include <memory>
#include <new>
#include <string_view>
#include <system_error>
#include <utility>

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <sqlite3.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820)
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include "Helpers.h"
#include "PathUtils.h"
#include "resource.h"

#pragma comment(lib, "Ole32.lib")

namespace ViewerSqliteEngine
{
namespace
{
using unique_sqlite_stmt   = std::unique_ptr<sqlite3_stmt, decltype(&sqlite3_finalize)>;
using unique_sqlite_backup = std::unique_ptr<sqlite3_backup, decltype(&sqlite3_backup_finish)>;

constexpr size_t kCopyChunkBytes            = 1024u * 1024u;
constexpr size_t kMaxCellChars              = 4096u;
constexpr size_t kMaxIdentifierChars        = 512u;
constexpr size_t kMaxExactIdentifierChars   = 16384u;
constexpr size_t kMaxRetainedTables         = 10000u;
constexpr uint64_t kMaxRetainedTableChars   = 4u * 1024u * 1024u;
constexpr uint32_t kMaxPageSize             = 1000u;
constexpr uint32_t kMaxQueryRowCap          = 100000u;
constexpr uint64_t kMaxVirtualSnapshotBytes = kMaxSnapshotBytes;
constexpr int kProgressOpcodeInterval        = 1000;
constexpr uint32_t kBackupTimeoutMs          = 60'000u;
constexpr unsigned int kTempPathAttempts     = 16u;
constexpr wchar_t kTempFilePrefix[]          = L"RedSalamander-ViewerSqlite-";

const int kViewerSqliteEngineModuleAnchor = 0;

struct TempSnapshot final
{
    TempSnapshot()                                    = default;
    TempSnapshot(const TempSnapshot&)                 = delete;
    TempSnapshot(TempSnapshot&&) noexcept             = default;
    TempSnapshot& operator=(const TempSnapshot&)      = delete;
    TempSnapshot& operator=(TempSnapshot&&) noexcept  = default;

    std::filesystem::path path;
    wil::unique_handle lifetimeHandle;
};

void ScavengeCrashedTempSnapshots(std::wstring_view tempDirectory) noexcept
{
    std::filesystem::path searchPattern(tempDirectory);
    searchPattern /= std::wstring(kTempFilePrefix) + L"*.sqlite";

    WIN32_FIND_DATAW findData{};
    wil::unique_hfind findHandle(FindFirstFileW(searchPattern.c_str(), &findData));
    if (! findHandle)
    {
        return;
    }

    do
    {
        if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
        {
            continue;
        }

        std::filesystem::path stalePath(tempDirectory);
        stalePath /= findData.cFileName;
        // A live snapshot's lifetime handle denies FILE_SHARE_DELETE, so this only
        // removes closed artifacts left behind by a terminated process.
        static_cast<void>(DeleteFileW(stalePath.c_str()));
    } while (FindNextFileW(findHandle.get(), &findData) != FALSE);
}

enum class WorkStopReason : uint8_t
{
    None,
    Cancelled,
    LimitExceeded,
};

struct SqliteProgressState final
{
    QueryCancellation cancellation;
    QueryWorkBudget budget;
    std::chrono::steady_clock::time_point startedAt = std::chrono::steady_clock::now();
    uint64_t observedVmSteps                         = 0;
    WorkStopReason stopReason                        = WorkStopReason::None;
};

class ScopedSqliteProgress final
{
public:
    ScopedSqliteProgress(sqlite3* db, SqliteProgressState& state) noexcept : _db(db)
    {
        if (_db != nullptr)
        {
            sqlite3_progress_handler(_db, kProgressOpcodeInterval, &ScopedSqliteProgress::OnProgress, &state);
        }
    }

    ScopedSqliteProgress(const ScopedSqliteProgress&)            = delete;
    ScopedSqliteProgress& operator=(const ScopedSqliteProgress&) = delete;

    ~ScopedSqliteProgress() noexcept
    {
        if (_db != nullptr)
        {
            sqlite3_progress_handler(_db, 0, nullptr, nullptr);
        }
    }

private:
    static int OnProgress(void* context) noexcept
    {
        auto* state = static_cast<SqliteProgressState*>(context);
        if (state == nullptr)
        {
            return 1;
        }

        state->observedVmSteps += static_cast<uint64_t>(kProgressOpcodeInterval);
        if (state->cancellation.IsCancellationRequested())
        {
            state->stopReason = WorkStopReason::Cancelled;
            return 1;
        }

        const bool stepLimitExceeded = state->budget.maxVmSteps != 0u && state->observedVmSteps > state->budget.maxVmSteps;
        const bool timeLimitExceeded = state->budget.maxElapsedMs != 0u &&
                                       std::chrono::steady_clock::now() - state->startedAt >=
                                           std::chrono::milliseconds(state->budget.maxElapsedMs);
        if (stepLimitExceeded || timeLimitExceeded)
        {
            state->stopReason = WorkStopReason::LimitExceeded;
            return 1;
        }

        return 0;
    }

    sqlite3* _db = nullptr;
};

class ConnectionUseCounter final
{
public:
    ConnectionUseCounter(std::atomic_uint32_t& active, std::atomic_uint32_t& maximum) noexcept : _active(active)
    {
        const uint32_t now = _active.fetch_add(1u, std::memory_order_acq_rel) + 1u;
        uint32_t currentMax = maximum.load(std::memory_order_relaxed);
        while (currentMax < now && ! maximum.compare_exchange_weak(currentMax, now, std::memory_order_release, std::memory_order_relaxed))
        {
        }
    }

    ConnectionUseCounter(const ConnectionUseCounter&)            = delete;
    ConnectionUseCounter& operator=(const ConnectionUseCounter&) = delete;

    ~ConnectionUseCounter() noexcept
    {
        static_cast<void>(_active.fetch_sub(1u, std::memory_order_acq_rel));
    }

private:
    std::atomic_uint32_t& _active;
};

[[nodiscard]] HINSTANCE GetEngineResourceInstance() noexcept
{
    HMODULE module = nullptr;
    const BOOL found = GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                          reinterpret_cast<LPCWSTR>(&kViewerSqliteEngineModuleAnchor),
                                          &module);
    return found != FALSE ? module : nullptr;
}

[[nodiscard]] std::wstring ReadEngineString(const UINT resourceId)
{
    return LoadStringResource(GetEngineResourceInstance(), resourceId);
}

template<typename... Args>
[[nodiscard]] std::wstring FormatEngineString(const UINT resourceId, Args&&... args)
{
    return FormatStringResource(GetEngineResourceInstance(), resourceId, std::forward<Args>(args)...);
}

[[nodiscard]] std::wstring SanitizeIdentifierText(std::wstring value) noexcept;

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text) noexcept
{
    size_t start = 0;
    while (start < text.size() && iswspace(text[start]) != 0)
    {
        ++start;
    }

    size_t end = text.size();
    while (end > start && iswspace(text[end - 1u]) != 0)
    {
        --end;
    }

    return std::wstring(text.substr(start, end - start));
}

[[nodiscard]] bool TailHasOnlyWhitespaceOrSemicolons(const char* tail) noexcept
{
    if (tail == nullptr)
    {
        return true;
    }

    const unsigned char* p = reinterpret_cast<const unsigned char*>(tail);
    while (*p != 0)
    {
        if (*p == ';' || std::isspace(*p) != 0)
        {
            ++p;
            continue;
        }

        if (p[0] == '-' && p[1] == '-')
        {
            p += 2;
            while (*p != 0 && *p != '\r' && *p != '\n')
            {
                ++p;
            }
            continue;
        }

        if (p[0] == '/' && p[1] == '*')
        {
            p += 2;
            bool closedBlock = false;
            while (*p != 0)
            {
                if (p[0] == '*' && p[1] == '/')
                {
                    p += 2;
                    closedBlock = true;
                    break;
                }
                ++p;
            }

            if (! closedBlock)
            {
                return false;
            }
            continue;
        }

        return false;
    }

    return true;
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    return Common::Strings::Utf8FromUtf16StrictOrEmpty(text);
}

[[nodiscard]] std::wstring Utf16FromBoundedUtf8String(const char* text, const size_t maxBytes, const bool appendEllipsis) noexcept
{
    if (text == nullptr || maxBytes == 0u)
    {
        return {};
    }

    const size_t measured = strnlen_s(text, maxBytes + 1u);
    const bool truncated  = measured > maxBytes;
    size_t bytesToConvert = std::min(measured, maxBytes);
    std::wstring converted;
    while (bytesToConvert != 0u)
    {
        converted = Utf16FromUtf8(std::string_view(text, bytesToConvert));
        if (! converted.empty())
        {
            break;
        }
        --bytesToConvert;
    }

    if (truncated && appendEllipsis)
    {
        converted.append(L"...");
    }
    return converted;
}

[[nodiscard]] std::wstring DescribeSqliteFailure(sqlite3* db, const int sqliteCode, std::wstring_view context) noexcept
{
    const char* raw     = (db != nullptr) ? sqlite3_errmsg(db) : sqlite3_errstr(sqliteCode);
    std::wstring detail = SanitizeIdentifierText(Utf16FromBoundedUtf8String(raw, 2048u, true));
    if (detail.empty())
    {
        detail = FormatEngineString(IDS_VIEWERSQLITE_ERROR_SQLITE_CODE_FMT, sqliteCode);
    }

    if (context.empty())
    {
        return detail;
    }

    return FormatEngineString(IDS_VIEWERSQLITE_ERROR_CONTEXT_FMT, context, detail);
}

[[nodiscard]] std::wstring DescribeSqliteFailure(sqlite3* db, const int sqliteCode, const UINT contextResourceId) noexcept
{
    return DescribeSqliteFailure(db, sqliteCode, ReadEngineString(contextResourceId));
}

[[nodiscard]] HRESULT DescribeStoppedWork(SqliteProgressState& progress,
                                          std::wstring& errorText,
                                          std::atomic_uint64_t* cancelledCount,
                                          std::atomic_uint64_t* workLimitCount) noexcept
{
    if (progress.stopReason == WorkStopReason::Cancelled || progress.cancellation.IsCancellationRequested())
    {
        progress.stopReason = WorkStopReason::Cancelled;
        if (cancelledCount != nullptr)
        {
            static_cast<void>(cancelledCount->fetch_add(1u, std::memory_order_relaxed));
        }
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_CANCELLED);
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    progress.stopReason = WorkStopReason::LimitExceeded;
    if (workLimitCount != nullptr)
    {
        static_cast<void>(workLimitCount->fetch_add(1u, std::memory_order_relaxed));
    }
    errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_WORK_LIMIT);
    return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
}

[[nodiscard]] HRESULT CheckCancellation(const QueryCancellation cancellation,
                                        std::wstring& errorText,
                                        std::atomic_uint64_t* cancelledCount = nullptr) noexcept
{
    if (! cancellation.IsCancellationRequested())
    {
        return S_OK;
    }

    if (cancelledCount != nullptr)
    {
        static_cast<void>(cancelledCount->fetch_add(1u, std::memory_order_relaxed));
    }
    errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_CANCELLED);
    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

[[nodiscard]] HRESULT OpenReadOnlyConnection(const std::filesystem::path& localPath, unique_sqlite3& db, std::wstring& errorText) noexcept
{
    const std::string pathUtf8 = Utf8FromUtf16(localPath.wstring());
    if (pathUtf8.empty())
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_DATABASE_PATH_UTF8);
        return E_INVALIDARG;
    }

    sqlite3* raw = nullptr;
    const int rc = sqlite3_open_v2(pathUtf8.c_str(), &raw, SQLITE_OPEN_READONLY | SQLITE_OPEN_NOMUTEX | SQLITE_OPEN_PRIVATECACHE, nullptr);
    db.reset(raw);
    if (rc != SQLITE_OK || ! db)
    {
        errorText = DescribeSqliteFailure(db.get(), rc, IDS_VIEWERSQLITE_ERROR_OPEN_DATABASE_CONTEXT);
        return E_FAIL;
    }

    static_cast<void>(sqlite3_extended_result_codes(db.get(), 1));
    static_cast<void>(sqlite3_db_config(db.get(), SQLITE_DBCONFIG_DEFENSIVE, 1, nullptr));
    static_cast<void>(sqlite3_db_config(db.get(), SQLITE_DBCONFIG_TRUSTED_SCHEMA, 0, nullptr));
    return S_OK;
}

[[nodiscard]] bool IsUnsafeIdentifierCodePoint(const wchar_t ch) noexcept
{
    const uint32_t value = static_cast<uint32_t>(ch);
    return value < 0x20u || (value >= 0x7Fu && value <= 0x9Fu) || value == 0x061Cu || value == 0x200Eu || value == 0x200Fu ||
           (value >= 0x202Au && value <= 0x202Eu) || (value >= 0x2066u && value <= 0x206Fu);
}

void TruncateUtf16Safely(std::wstring& value, const size_t maxChars)
{
    if (value.size() <= maxChars)
    {
        return;
    }

    size_t truncatedLength = maxChars;
    if (truncatedLength > 0u && truncatedLength < value.size())
    {
        const wchar_t trailing         = value[truncatedLength - 1u];
        const wchar_t leading          = value[truncatedLength];
        const bool splitsSurrogatePair = (trailing >= static_cast<wchar_t>(0xD800) && trailing <= static_cast<wchar_t>(0xDBFF)) &&
                                         (leading >= static_cast<wchar_t>(0xDC00) && leading <= static_cast<wchar_t>(0xDFFF));
        if (splitsSurrogatePair)
        {
            --truncatedLength;
        }
    }

    value.resize(truncatedLength);
    value.append(L"...");
}

[[nodiscard]] std::wstring SanitizeCellText(std::wstring value) noexcept
{
    for (wchar_t& ch : value)
    {
        switch (ch)
        {
            case L'\0': ch = L'?'; break;
            case L'\r':
            case L'\n':
            case L'\t': ch = L' '; break;
            default: break;
        }
    }

    TruncateUtf16Safely(value, kMaxCellChars);
    return value;
}

[[nodiscard]] std::wstring SanitizeIdentifierText(std::wstring value) noexcept
{
    for (wchar_t& ch : value)
    {
        if (IsUnsafeIdentifierCodePoint(ch))
        {
            ch = L' ';
        }
    }

    TruncateUtf16Safely(value, kMaxIdentifierChars);
    return value;
}

[[nodiscard]] std::wstring ReadBoundedRawTextValue(sqlite3_stmt* stmt, const int columnIndex, const size_t maxChars) noexcept
{
    const void* text16 = sqlite3_column_text16(stmt, columnIndex);
    const int bytes16  = sqlite3_column_bytes16(stmt, columnIndex);
    if (text16 == nullptr || bytes16 <= 0)
    {
        return {};
    }

    const auto length = static_cast<size_t>(bytes16 / static_cast<int>(sizeof(wchar_t)));
    const auto* wide  = static_cast<const wchar_t*>(text16);
    const size_t copiedLength = std::min(length, maxChars + 1u);
    std::wstring value(wide, copiedLength);
    TruncateUtf16Safely(value, maxChars);
    return value;
}

[[nodiscard]] std::wstring ReadCellValue(sqlite3_stmt* stmt, const int columnIndex) noexcept
{
    const int type = sqlite3_column_type(stmt, columnIndex);
    switch (type)
    {
        case SQLITE_NULL: return ReadEngineString(IDS_VIEWERSQLITE_VALUE_NULL);
        case SQLITE_BLOB:
        {
            const int bytes = sqlite3_column_bytes(stmt, columnIndex);
            return FormatEngineString(IDS_VIEWERSQLITE_VALUE_BLOB_FMT, bytes);
        }
        case SQLITE_INTEGER:
        case SQLITE_FLOAT:
        {
            const unsigned char* text = sqlite3_column_text(stmt, columnIndex);
            if (text == nullptr)
            {
                return {};
            }
            const int bytes = sqlite3_column_bytes(stmt, columnIndex);
            return bytes > 0 ? Utf16FromUtf8(std::string_view(reinterpret_cast<const char*>(text), static_cast<size_t>(bytes))) : std::wstring{};
        }
        case SQLITE_TEXT:
        default: return SanitizeCellText(ReadBoundedRawTextValue(stmt, columnIndex, kMaxCellChars));
    }
}

[[nodiscard]] HRESULT ReadQueryPage(sqlite3* db,
                                    unique_sqlite_stmt& stmt,
                                    std::wstring executedSql,
                                    const uint32_t rowLimit,
                                    const uint64_t rowOffset,
                                    const bool markTruncatedOnOverflow,
                                    const QueryWorkBudget budget,
                                    SqliteProgressState& progress,
                                    QueryPage& page,
                                    std::wstring& errorText,
                                    std::atomic_uint64_t* cancelledCount,
                                    std::atomic_uint64_t* workLimitCount) noexcept
{
    page             = {};
    page.executedSql = std::move(executedSql);
    page.rowOffset   = rowOffset;

    const int columnCount = sqlite3_column_count(stmt.get());
    page.columns.reserve(static_cast<size_t>(std::max(columnCount, 0)));
    for (int columnIndex = 0; columnIndex < columnCount; ++columnIndex)
    {
        const char* rawName = sqlite3_column_name(stmt.get(), columnIndex);
        const char* rawDecl = sqlite3_column_decltype(stmt.get(), columnIndex);
        ColumnInfo columnInfo{};
        columnInfo.name = SanitizeIdentifierText(Utf16FromBoundedUtf8String(rawName, kMaxIdentifierChars * 4u, true));
        columnInfo.declaredType = SanitizeIdentifierText(Utf16FromBoundedUtf8String(rawDecl, kMaxIdentifierChars * 4u, true));
        page.columns.push_back(std::move(columnInfo));
    }

    const uint32_t effectiveLimit = std::clamp<uint32_t>(rowLimit, 1u, kMaxQueryRowCap);
    const uint64_t requestedCells = static_cast<uint64_t>(effectiveLimit) * static_cast<uint64_t>(std::max(columnCount, 0));
    if (budget.maxResultCells != 0u && requestedCells > budget.maxResultCells)
    {
        if (workLimitCount != nullptr)
        {
            static_cast<void>(workLimitCount->fetch_add(1u, std::memory_order_relaxed));
        }
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_RESULT_LIMIT);
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }

    uint64_t materializedChars = 0;
    for (;;)
    {
        const int rc = sqlite3_step(stmt.get());
        if (rc == SQLITE_DONE)
        {
            break;
        }

        if (rc != SQLITE_ROW)
        {
            if (rc == SQLITE_INTERRUPT || progress.stopReason != WorkStopReason::None || progress.cancellation.IsCancellationRequested())
            {
                return DescribeStoppedWork(progress, errorText, cancelledCount, workLimitCount);
            }
            errorText = DescribeSqliteFailure(db, rc, IDS_VIEWERSQLITE_ERROR_READ_QUERY_ROWS_CONTEXT);
            return E_FAIL;
        }

        if (page.rows.size() >= effectiveLimit)
        {
            page.hasMore   = true;
            page.truncated = markTruncatedOnOverflow;
            break;
        }

        std::vector<std::wstring> row;
        row.reserve(static_cast<size_t>(std::max(columnCount, 0)));
        for (int columnIndex = 0; columnIndex < columnCount; ++columnIndex)
        {
            std::wstring value = ReadCellValue(stmt.get(), columnIndex);
            const uint64_t valueChars = static_cast<uint64_t>(value.size());
            if (valueChars > std::numeric_limits<uint64_t>::max() - materializedChars ||
                (budget.maxResultChars != 0u && materializedChars + valueChars > budget.maxResultChars))
            {
                if (workLimitCount != nullptr)
                {
                    static_cast<void>(workLimitCount->fetch_add(1u, std::memory_order_relaxed));
                }
                errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_RESULT_LIMIT);
                return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
            }
            materializedChars += valueChars;
            row.push_back(std::move(value));
        }

        page.rows.push_back(std::move(row));
    }

    return S_OK;
}

[[nodiscard]] HRESULT PrepareSingleReadonlyStatement(sqlite3* db,
                                                     std::wstring_view sql,
                                                     unique_sqlite_stmt& stmt,
                                                     std::wstring& executedSql,
                                                     std::wstring& errorText,
                                                     SqliteProgressState& progress,
                                                     std::atomic_uint64_t* cancelledCount,
                                                     std::atomic_uint64_t* workLimitCount) noexcept
{
    executedSql = TrimWhitespace(sql);
    if (executedSql.empty())
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_STATUS_ENTER_QUERY);
        return E_INVALIDARG;
    }

    const std::string sqlUtf8 = Utf8FromUtf16(executedSql);
    if (sqlUtf8.empty())
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SQL_UTF8);
        return E_INVALIDARG;
    }

    sqlite3_stmt* raw = nullptr;
    const char* tail  = nullptr;
    const int rc      = sqlite3_prepare_v2(db, sqlUtf8.c_str(), static_cast<int>(sqlUtf8.size()), &raw, &tail);
    stmt.reset(raw);
    if (rc != SQLITE_OK)
    {
        if (rc == SQLITE_INTERRUPT || progress.stopReason != WorkStopReason::None || progress.cancellation.IsCancellationRequested())
        {
            return DescribeStoppedWork(progress, errorText, cancelledCount, workLimitCount);
        }
        errorText = DescribeSqliteFailure(db, rc, IDS_VIEWERSQLITE_ERROR_PREPARE_QUERY_CONTEXT);
        return E_FAIL;
    }

    if (! stmt)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_STATUS_ENTER_QUERY);
        return E_INVALIDARG;
    }

    if (! TailHasOnlyWhitespaceOrSemicolons(tail))
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SINGLE_STATEMENT);
        return E_INVALIDARG;
    }

    if (sqlite3_stmt_readonly(stmt.get()) == 0)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_READ_ONLY_STATEMENT);
        return E_ACCESSDENIED;
    }

    return S_OK;
}

[[nodiscard]] bool IsBuiltinFileSystem(IFileSystem* fileSystem) noexcept
{
    if (fileSystem == nullptr)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    const HRESULT infoHr = fileSystem->QueryInterface(__uuidof(IInformations), info.put_void());
    if (FAILED(infoHr) || ! info)
    {
        return false;
    }

    const PluginMetaData* meta = nullptr;
    const HRESULT metaHr       = info->GetMetaData(&meta);
    if (FAILED(metaHr) || meta == nullptr || meta->id == nullptr)
    {
        return false;
    }

    return _wcsicmp(meta->id, L"builtin/file-system") == 0;
}

[[nodiscard]] HRESULT CreateCrashCleanedTempSnapshot(TempSnapshot& snapshot, std::wstring& errorText) noexcept
{
    snapshot = {};

    std::array<wchar_t, 32768> tempDirectory{};
    const DWORD tempChars = GetTempPathW(static_cast<DWORD>(tempDirectory.size()), tempDirectory.data());
    if (tempChars == 0u || tempChars >= tempDirectory.size())
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_TEMP_DIRECTORY);
        return HRESULT_FROM_WIN32(GetLastError());
    }

    ScavengeCrashedTempSnapshots(tempDirectory.data());

    std::wstring tempPath;
    wil::unique_handle file;
    const Common::Paths::UniqueSiblingFileOptions options{.prefix             = kTempFilePrefix,
                                                           .suffix             = L".sqlite",
                                                           .desiredAccess      = GENERIC_READ | GENERIC_WRITE,
                                                           .shareMode          = FILE_SHARE_READ | FILE_SHARE_WRITE,
                                                           .flagsAndAttributes = FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_SEQUENTIAL_SCAN,
                                                           .maximumAttempts   = kTempPathAttempts};
    const HRESULT createHr = Common::Paths::CreateUniqueFileInDirectory(tempDirectory.data(), options, tempPath, file);
    if (FAILED(createHr))
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_TEMP_FILE);
        return createHr;
    }
    snapshot.path           = std::filesystem::path(std::move(tempPath));
    snapshot.lifetimeHandle = std::move(file);
    return S_OK;
}

[[nodiscard]] HRESULT RewindAndTruncateSnapshot(HANDLE file, std::wstring& errorText) noexcept
{
    LARGE_INTEGER beginning{};
    if (SetFilePointerEx(file, beginning, nullptr, FILE_BEGIN) == 0 || SetEndOfFile(file) == 0)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SNAPSHOT_CREATE);
        return HRESULT_FROM_WIN32(GetLastError());
    }
    return S_OK;
}

[[nodiscard]] HRESULT GetSnapshotSize(HANDLE file, uint64_t& sizeBytes, std::wstring& errorText) noexcept
{
    LARGE_INTEGER size{};
    if (GetFileSizeEx(file, &size) == 0 || size.QuadPart < 0)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SOURCE_SIZE);
        return HRESULT_FROM_WIN32(GetLastError());
    }
    sizeBytes = static_cast<uint64_t>(size.QuadPart);
    return S_OK;
}

[[nodiscard]] HRESULT ReadSqlitePragmaUint64(sqlite3* db,
                                             const char* sql,
                                             uint64_t& value,
                                             std::wstring& errorText) noexcept
{
    value = 0u;
    sqlite3_stmt* raw = nullptr;
    const int prepareRc = sqlite3_prepare_v2(db, sql, -1, &raw, nullptr);
    unique_sqlite_stmt stmt(raw, sqlite3_finalize);
    if (prepareRc != SQLITE_OK || ! stmt)
    {
        errorText = DescribeSqliteFailure(db, prepareRc, IDS_VIEWERSQLITE_ERROR_OPEN_DATABASE_CONTEXT);
        return E_FAIL;
    }

    const int stepRc = sqlite3_step(stmt.get());
    if (stepRc != SQLITE_ROW)
    {
        errorText = DescribeSqliteFailure(db, stepRc, IDS_VIEWERSQLITE_ERROR_OPEN_DATABASE_CONTEXT);
        return E_FAIL;
    }

    const sqlite3_int64 signedValue = sqlite3_column_int64(stmt.get(), 0);
    if (signedValue < 0)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_LOCAL_SNAPSHOT_TOO_LARGE);
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }

    value = static_cast<uint64_t>(signedValue);
    return S_OK;
}

[[nodiscard]] HRESULT BackupLocalDatabase(const std::filesystem::path& sourcePath,
                                          const TempSnapshot& snapshot,
                                          const QueryCancellation cancellation,
                                          const uint64_t maxSnapshotBytes,
                                          uint64_t& snapshotBytes,
                                          std::wstring& errorText) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    HRESULT hr           = CheckCancellation(cancellation, errorText);
    if (FAILED(hr))
    {
        Debug::Perf::Emit(L"viewer.sqlite.snapshot_us", L"local-backup-cancelled", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, hr);
        return hr;
    }

    hr = RewindAndTruncateSnapshot(snapshot.lifetimeHandle.get(), errorText);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::string sourceUtf8   = Utf8FromUtf16(sourcePath.wstring());
    const std::string snapshotUtf8 = Utf8FromUtf16(snapshot.path.wstring());
    if (sourceUtf8.empty() || snapshotUtf8.empty())
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_DATABASE_PATH_UTF8);
        return E_INVALIDARG;
    }

    sqlite3* rawSource = nullptr;
    int rc = sqlite3_open_v2(sourceUtf8.c_str(), &rawSource, SQLITE_OPEN_READONLY | SQLITE_OPEN_NOMUTEX | SQLITE_OPEN_PRIVATECACHE, nullptr);
    unique_sqlite3 source(rawSource);
    if (rc != SQLITE_OK || ! source)
    {
        errorText = DescribeSqliteFailure(source.get(), rc, IDS_VIEWERSQLITE_ERROR_OPEN_DATABASE_CONTEXT);
        hr        = E_FAIL;
        Debug::Perf::Emit(L"viewer.sqlite.snapshot_us", L"local-backup-open-source", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, hr);
        return hr;
    }

    rc = sqlite3_exec(source.get(), "BEGIN", nullptr, nullptr, nullptr);
    if (rc != SQLITE_OK)
    {
        errorText = DescribeSqliteFailure(source.get(), rc, IDS_VIEWERSQLITE_ERROR_OPEN_DATABASE_CONTEXT);
        return E_FAIL;
    }
    const auto endSourceReadTransaction = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(sqlite3_exec(source.get(), "ROLLBACK", nullptr, nullptr, nullptr));
    });

    uint64_t pageCount = 0u;
    uint64_t pageSize  = 0u;
    hr = ReadSqlitePragmaUint64(source.get(), "PRAGMA page_count", pageCount, errorText);
    if (SUCCEEDED(hr))
    {
        hr = ReadSqlitePragmaUint64(source.get(), "PRAGMA page_size", pageSize, errorText);
    }
    if (FAILED(hr))
    {
        return hr;
    }
    if (pageSize == 0u || maxSnapshotBytes == 0u || pageCount > maxSnapshotBytes / pageSize)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_LOCAL_SNAPSHOT_TOO_LARGE);
        hr        = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        Debug::Perf::Emit(L"viewer.sqlite.snapshot_us", L"local-backup-size-rejected", Debug::Perf::ElapsedUs(startedAt), pageCount, pageSize, hr);
        return hr;
    }
    const uint64_t preflightBytes = pageCount * pageSize;

    sqlite3* rawDestination = nullptr;
    rc = sqlite3_open_v2(snapshotUtf8.c_str(),
                         &rawDestination,
                         SQLITE_OPEN_READWRITE | SQLITE_OPEN_NOMUTEX | SQLITE_OPEN_PRIVATECACHE,
                         nullptr);
    unique_sqlite3 destination(rawDestination);
    if (rc != SQLITE_OK || ! destination)
    {
        errorText = DescribeSqliteFailure(destination.get(), rc, IDS_VIEWERSQLITE_ERROR_BACKUP_DESTINATION_CONTEXT);
        hr        = E_FAIL;
        Debug::Perf::Emit(L"viewer.sqlite.snapshot_us", L"local-backup-open-destination", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, hr);
        return hr;
    }

    sqlite3_backup* rawBackup = sqlite3_backup_init(destination.get(), "main", source.get(), "main");
    unique_sqlite_backup backup(rawBackup, sqlite3_backup_finish);
    if (! backup)
    {
        errorText = DescribeSqliteFailure(destination.get(), sqlite3_errcode(destination.get()), IDS_VIEWERSQLITE_ERROR_BACKUP_INITIALIZE_CONTEXT);
        hr        = E_FAIL;
        Debug::Perf::Emit(L"viewer.sqlite.snapshot_us", L"local-backup-init", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, hr);
        return hr;
    }

    for (;;)
    {
        if (cancellation.IsCancellationRequested())
        {
            errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_CANCELLED);
            hr        = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            break;
        }
        if (std::chrono::steady_clock::now() - startedAt >= std::chrono::milliseconds(kBackupTimeoutMs))
        {
            errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_WORK_LIMIT);
            hr        = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
            break;
        }

        rc = sqlite3_backup_step(backup.get(), 256);
        const int observedPageCount = sqlite3_backup_pagecount(backup.get());
        if (observedPageCount < 0 || static_cast<uint64_t>(observedPageCount) > maxSnapshotBytes / pageSize)
        {
            errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_LOCAL_SNAPSHOT_TOO_LARGE);
            hr        = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
            break;
        }
        if (rc == SQLITE_DONE)
        {
            hr = S_OK;
            break;
        }
        if (rc == SQLITE_OK)
        {
            continue;
        }
        if (rc == SQLITE_BUSY || rc == SQLITE_LOCKED)
        {
            static_cast<void>(sqlite3_sleep(2));
            continue;
        }

        errorText = DescribeSqliteFailure(destination.get(), rc, IDS_VIEWERSQLITE_ERROR_BACKUP_STEP_CONTEXT);
        hr        = E_FAIL;
        break;
    }

    const int finishRc = sqlite3_backup_finish(backup.release());
    if (SUCCEEDED(hr) && finishRc != SQLITE_OK)
    {
        errorText = DescribeSqliteFailure(destination.get(), finishRc, IDS_VIEWERSQLITE_ERROR_BACKUP_STEP_CONTEXT);
        hr        = E_FAIL;
    }

    destination.reset();
    if (SUCCEEDED(hr))
    {
        hr = GetSnapshotSize(snapshot.lifetimeHandle.get(), snapshotBytes, errorText);
        if (SUCCEEDED(hr) && snapshotBytes > maxSnapshotBytes)
        {
            errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_LOCAL_SNAPSHOT_TOO_LARGE);
            hr        = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }
    }

    Debug::Perf::Emit(L"viewer.sqlite.snapshot_us",
                      SUCCEEDED(hr) ? L"local-backup" : L"local-backup-failed",
                      Debug::Perf::ElapsedUs(startedAt),
                      snapshotBytes,
                      preflightBytes,
                      hr);
    return hr;
}

[[nodiscard]] HRESULT CopyReaderToSnapshot(IFileReader* reader,
                                           const TempSnapshot& snapshot,
                                           const QueryCancellation cancellation,
                                           uint64_t& snapshotBytes,
                                           std::wstring& errorText) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    if (reader == nullptr)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_READER_UNAVAILABLE);
        return E_INVALIDARG;
    }

    uint64_t advertisedBefore = 0;
    HRESULT hr = reader->GetSize(&advertisedBefore);
    if (FAILED(hr))
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SOURCE_SIZE);
        return hr;
    }
    if (advertisedBefore > kMaxVirtualSnapshotBytes)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SOURCE_TOO_LARGE);
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }

    hr = RewindAndTruncateSnapshot(snapshot.lifetimeHandle.get(), errorText);
    if (FAILED(hr))
    {
        return hr;
    }

    uint64_t rewindPosition = 0;
    hr                      = reader->Seek(0, FILE_BEGIN, &rewindPosition);
    if (FAILED(hr) || rewindPosition != 0u)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SOURCE_SEEK);
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_SEEK);
    }

    auto buffer = std::unique_ptr<BYTE[]>(new (std::nothrow) BYTE[kCopyChunkBytes]);
    if (! buffer)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SNAPSHOT_OOM);
        return E_OUTOFMEMORY;
    }

    uint64_t copiedBytes = 0;
    std::array<BYTE, 20> databaseHeader{};
    size_t databaseHeaderBytes = 0u;
    bool databaseHeaderValidated = false;
    for (;;)
    {
        hr = CheckCancellation(cancellation, errorText);
        if (FAILED(hr))
        {
            break;
        }

        unsigned long readBytes = 0;
        hr = reader->Read(buffer.get(), static_cast<unsigned long>(kCopyChunkBytes), &readBytes);
        if (FAILED(hr))
        {
            errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SOURCE_READ);
            break;
        }

        if (readBytes == 0)
        {
            break;
        }

        if (readBytes > static_cast<unsigned long>(kCopyChunkBytes))
        {
            errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SOURCE_READ);
            hr        = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            break;
        }

        if (readBytes > kMaxVirtualSnapshotBytes - copiedBytes || copiedBytes + readBytes > advertisedBefore)
        {
            errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SOURCE_CHANGED);
            hr        = HRESULT_FROM_WIN32(ERROR_FILE_INVALID);
            break;
        }

        if (! databaseHeaderValidated && databaseHeaderBytes < databaseHeader.size())
        {
            const size_t headerRemaining = databaseHeader.size() - databaseHeaderBytes;
            const size_t headerCopyBytes = std::min<size_t>(headerRemaining, readBytes);
            std::memcpy(databaseHeader.data() + databaseHeaderBytes, buffer.get(), headerCopyBytes);
            databaseHeaderBytes += headerCopyBytes;
            if (databaseHeaderBytes == databaseHeader.size())
            {
                static constexpr std::array<BYTE, 16> kSqliteHeader{{'S', 'Q', 'L', 'i', 't', 'e', ' ', 'f', 'o', 'r', 'm', 'a', 't', ' ', '3', 0}};
                databaseHeaderValidated = true;
                if (std::equal(kSqliteHeader.begin(), kSqliteHeader.end(), databaseHeader.begin()) &&
                    (databaseHeader[18] == 2u || databaseHeader[19] == 2u))
                {
                    errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_VIRTUAL_WAL_UNSUPPORTED);
                    hr        = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                    break;
                }
            }
        }

        DWORD writtenBytes = 0;
        const BOOL writeOk  = WriteFile(snapshot.lifetimeHandle.get(), buffer.get(), readBytes, &writtenBytes, nullptr);
        if (writeOk == 0 || writtenBytes != readBytes)
        {
            errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SNAPSHOT_WRITE);
            hr        = HRESULT_FROM_WIN32(writeOk == 0 ? GetLastError() : ERROR_WRITE_FAULT);
            break;
        }
        copiedBytes += readBytes;
    }

    uint64_t advertisedAfter = 0;
    if (SUCCEEDED(hr))
    {
        hr = reader->GetSize(&advertisedAfter);
        if (FAILED(hr))
        {
            errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SOURCE_SIZE);
        }
        else if (advertisedBefore != advertisedAfter || copiedBytes != advertisedBefore)
        {
            errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SOURCE_CHANGED);
            hr        = HRESULT_FROM_WIN32(ERROR_FILE_INVALID);
        }
    }

    if (SUCCEEDED(hr) && FlushFileBuffers(snapshot.lifetimeHandle.get()) == 0)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SNAPSHOT_FLUSH);
        hr        = HRESULT_FROM_WIN32(GetLastError());
    }

    snapshotBytes = copiedBytes;
    Debug::Perf::Emit(L"viewer.sqlite.snapshot_us",
                      SUCCEEDED(hr) ? L"virtual-copy" : L"virtual-copy-rejected",
                      Debug::Perf::ElapsedUs(startedAt),
                      copiedBytes,
                      advertisedBefore,
                      hr);
    return hr;
}

[[nodiscard]] HRESULT ResolveTableColumnCount(sqlite3* db,
                                              std::wstring_view tableName,
                                              size_t& columnCount,
                                              std::wstring& errorText,
                                              SqliteProgressState& progress,
                                              std::atomic_uint64_t* cancelledCount,
                                              std::atomic_uint64_t* workLimitCount) noexcept
{
    const std::wstring sql = std::format(L"SELECT * FROM {} LIMIT 0", QuoteIdentifier(tableName));
    const std::string sqlUtf8 = Utf8FromUtf16(sql);
    if (sqlUtf8.empty())
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_PREVIEW_SQL);
        return E_INVALIDARG;
    }

    sqlite3_stmt* raw = nullptr;
    const int rc = sqlite3_prepare_v2(db, sqlUtf8.c_str(), static_cast<int>(sqlUtf8.size()), &raw, nullptr);
    unique_sqlite_stmt stmt(raw, sqlite3_finalize);
    if (rc != SQLITE_OK || ! stmt)
    {
        if (rc == SQLITE_INTERRUPT || progress.stopReason != WorkStopReason::None || progress.cancellation.IsCancellationRequested())
        {
            return DescribeStoppedWork(progress, errorText, cancelledCount, workLimitCount);
        }
        const std::wstring context = FormatEngineString(IDS_VIEWERSQLITE_ERROR_TABLE_METADATA_FMT, SanitizeIdentifierText(std::wstring(tableName)));
        errorText                 = DescribeSqliteFailure(db, rc, context);
        return E_FAIL;
    }

    columnCount = static_cast<size_t>(std::max(sqlite3_column_count(stmt.get()), 0));
    return S_OK;
}
} // namespace

void SqliteConnectionDeleter::operator()(sqlite3* db) const noexcept
{
    if (db != nullptr)
    {
        static_cast<void>(sqlite3_close_v2(db));
    }
}

DatabaseSource::DatabaseSource(std::filesystem::path localPath,
                               std::wstring displayName,
                               wil::unique_handle snapshotLifetimeHandle,
                               unique_sqlite3 connection,
                               const SnapshotKind snapshotKind,
                               const uint64_t snapshotBytes) noexcept
    : _localPath(std::move(localPath)),
      _displayName(std::move(displayName)),
      _snapshotLifetimeHandle(std::move(snapshotLifetimeHandle)),
      _connection(std::move(connection)),
      _snapshotKind(snapshotKind),
      _snapshotBytes(snapshotBytes)
{
}

DatabaseSource::~DatabaseSource() noexcept
{
    _connection.reset();
    _snapshotLifetimeHandle.reset();

    std::error_code ec;
    std::filesystem::remove(_localPath, ec);
}

const std::filesystem::path& DatabaseSource::GetLocalPath() const noexcept
{
    return _localPath;
}

const std::wstring& DatabaseSource::GetDisplayName() const noexcept
{
    return _displayName;
}

SourceDebugSnapshot DatabaseSource::GetDebugSnapshot() const noexcept
{
    SourceDebugSnapshot snapshot{};
    snapshot.cachedConnectionOpenCount = _connection ? 1u : 0u;
    snapshot.operationCount             = _operationCount.load(std::memory_order_relaxed);
    snapshot.cancelledOperationCount    = _cancelledOperationCount.load(std::memory_order_relaxed);
    snapshot.workLimitFailureCount      = _workLimitFailureCount.load(std::memory_order_relaxed);
    snapshot.maxConcurrentConnectionUse = _maxConcurrentConnectionUse.load(std::memory_order_relaxed);
    snapshot.snapshotBytes              = _snapshotBytes;
    snapshot.snapshotKind               = _snapshotKind;
    return snapshot;
}

HRESULT DatabaseSource::ListTables(std::vector<TableInfo>& tables,
                                   std::wstring& errorText,
                                   const QueryCancellation cancellation,
                                   const QueryWorkBudget budget) const noexcept
{
    tables.clear();
    errorText.clear();

    HRESULT hr = CheckCancellation(cancellation, errorText, &_cancelledOperationCount);
    if (FAILED(hr))
    {
        return hr;
    }

    std::scoped_lock lock(_connectionMutex);
    ConnectionUseCounter connectionUse(_activeConnectionUse, _maxConcurrentConnectionUse);
    hr = CheckCancellation(cancellation, errorText, &_cancelledOperationCount);
    if (FAILED(hr))
    {
        return hr;
    }
    return ListTablesLocked(tables, errorText, cancellation, budget);
}

HRESULT DatabaseSource::ListTablesLocked(std::vector<TableInfo>& tables,
                                         std::wstring& errorText,
                                         const QueryCancellation cancellation,
                                         const QueryWorkBudget budget) const noexcept
{
    static_cast<void>(_operationCount.fetch_add(1u, std::memory_order_relaxed));
    if (! _connection)
    {
        errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_CONNECTION_UNAVAILABLE);
        return E_UNEXPECTED;
    }

    constexpr char kSql[] = "SELECT name, type "
                            "FROM sqlite_schema "
                            "WHERE type IN ('table', 'view') "
                            "  AND name NOT LIKE 'sqlite_%' "
                            "ORDER BY CASE type WHEN 'table' THEN 0 ELSE 1 END, name COLLATE NOCASE";

    SqliteProgressState progress{.cancellation = cancellation, .budget = budget};
    ScopedSqliteProgress progressHandler(_connection.get(), progress);

    sqlite3_stmt* raw = nullptr;
    const int rc      = sqlite3_prepare_v2(_connection.get(), kSql, -1, &raw, nullptr);
    unique_sqlite_stmt stmt(raw, sqlite3_finalize);
    if (rc != SQLITE_OK || ! stmt)
    {
        if (rc == SQLITE_INTERRUPT || progress.stopReason != WorkStopReason::None || cancellation.IsCancellationRequested())
        {
            return DescribeStoppedWork(progress, errorText, &_cancelledOperationCount, &_workLimitFailureCount);
        }
        errorText = DescribeSqliteFailure(_connection.get(), rc, IDS_VIEWERSQLITE_ERROR_ENUMERATE_TABLES_CONTEXT);
        return E_FAIL;
    }

    _tableColumnCounts.clear();
    uint64_t retainedTableChars = 0u;
    for (;;)
    {
        const int stepRc = sqlite3_step(stmt.get());
        if (stepRc == SQLITE_DONE)
        {
            break;
        }
        if (stepRc != SQLITE_ROW)
        {
            if (stepRc == SQLITE_INTERRUPT || progress.stopReason != WorkStopReason::None || cancellation.IsCancellationRequested())
            {
                return DescribeStoppedWork(progress, errorText, &_cancelledOperationCount, &_workLimitFailureCount);
            }
            errorText = DescribeSqliteFailure(_connection.get(), stepRc, IDS_VIEWERSQLITE_ERROR_ENUMERATE_TABLES_CONTEXT);
            return E_FAIL;
        }

        TableInfo table{};
        static_cast<void>(sqlite3_column_text16(stmt.get(), 0));
        const int tableNameBytes = sqlite3_column_bytes16(stmt.get(), 0);
        if (tableNameBytes < 0 || static_cast<uint64_t>(tableNameBytes) > static_cast<uint64_t>(kMaxExactIdentifierChars * sizeof(wchar_t)))
        {
            errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_IDENTIFIER_TOO_LONG);
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }

        table.name        = ReadBoundedRawTextValue(stmt.get(), 0, kMaxExactIdentifierChars);
        table.displayName = SanitizeIdentifierText(table.name);
        table.kind        = SanitizeIdentifierText(ReadBoundedRawTextValue(stmt.get(), 1, kMaxIdentifierChars));
        if (! table.name.empty())
        {
            const uint64_t tableChars = static_cast<uint64_t>(table.name.size() + table.displayName.size() + table.kind.size());
            if (tables.size() >= kMaxRetainedTables || tableChars > std::numeric_limits<uint64_t>::max() - retainedTableChars ||
                retainedTableChars + tableChars > kMaxRetainedTableChars)
            {
                errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_RESULT_LIMIT);
                static_cast<void>(_workLimitFailureCount.fetch_add(1u, std::memory_order_relaxed));
                return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
            }
            retainedTableChars += tableChars;
            _tableColumnCounts.try_emplace(table.name, kNoSortColumn);
            tables.push_back(std::move(table));
        }
    }

    return S_OK;
}

QueryPageResult DatabaseSource::LoadTablePage(std::wstring_view tableName,
                                              const uint32_t pageSize,
                                              const uint64_t rowOffset,
                                              const size_t orderByColumnIndex,
                                              const TableSortDirection sortDirection,
                                              const QueryCancellation cancellation,
                                              const QueryWorkBudget budget) const noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    QueryPageResult result{};
    const std::wstring exactTable(tableName);
    if (exactTable.empty())
    {
        result.hr        = E_INVALIDARG;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_STATUS_SELECT_TABLE);
        return result;
    }
    if (rowOffset > static_cast<uint64_t>(std::numeric_limits<int64_t>::max()))
    {
        result.hr        = E_INVALIDARG;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_ROW_OFFSET);
        return result;
    }

    result.hr = CheckCancellation(cancellation, result.errorText, &_cancelledOperationCount);
    if (FAILED(result.hr))
    {
        return result;
    }

    std::scoped_lock lock(_connectionMutex);
    ConnectionUseCounter connectionUse(_activeConnectionUse, _maxConcurrentConnectionUse);
    static_cast<void>(_operationCount.fetch_add(1u, std::memory_order_relaxed));
    if (! _connection)
    {
        result.hr        = E_UNEXPECTED;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_CONNECTION_UNAVAILABLE);
        return result;
    }

    result.hr = CheckCancellation(cancellation, result.errorText, &_cancelledOperationCount);
    if (FAILED(result.hr))
    {
        return result;
    }

    SqliteProgressState progress{.cancellation = cancellation, .budget = budget};
    ScopedSqliteProgress progressHandler(_connection.get(), progress);

    auto tableIt = _tableColumnCounts.find(exactTable);
    if (tableIt == _tableColumnCounts.end())
    {
        result.hr        = E_INVALIDARG;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_STATUS_SELECT_TABLE);
        return result;
    }

    if (sortDirection != TableSortDirection::None && sortDirection != TableSortDirection::Ascending && sortDirection != TableSortDirection::Descending)
    {
        result.hr        = E_INVALIDARG;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SORT_COLUMN);
        return result;
    }

    size_t columnCount = tableIt->second;
    if (columnCount == kNoSortColumn)
    {
        result.hr = ResolveTableColumnCount(_connection.get(),
                                            exactTable,
                                            columnCount,
                                            result.errorText,
                                            progress,
                                            &_cancelledOperationCount,
                                            &_workLimitFailureCount);
        if (FAILED(result.hr))
        {
            return result;
        }
        tableIt->second = columnCount;
    }

    if (sortDirection != TableSortDirection::None && (orderByColumnIndex == kNoSortColumn || orderByColumnIndex >= columnCount))
    {
        result.hr        = E_INVALIDARG;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SORT_COLUMN);
        return result;
    }

    const uint32_t requestedPageSize = std::clamp<uint32_t>(pageSize, 1u, kMaxPageSize);
    const uint32_t fetchPageSize     = requestedPageSize + 1u;
    const std::wstring sql           = BuildTablePreviewSql(exactTable, fetchPageSize, rowOffset, orderByColumnIndex, sortDirection);
    const std::string sqlUtf8        = Utf8FromUtf16(sql);
    if (sqlUtf8.empty())
    {
        result.hr        = E_INVALIDARG;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_PREVIEW_SQL);
        return result;
    }

    sqlite3_stmt* raw = nullptr;
    const int rc      = sqlite3_prepare_v2(_connection.get(), sqlUtf8.c_str(), static_cast<int>(sqlUtf8.size()), &raw, nullptr);
    unique_sqlite_stmt stmt(raw, sqlite3_finalize);
    if (rc != SQLITE_OK || ! stmt)
    {
        if (rc == SQLITE_INTERRUPT || progress.stopReason != WorkStopReason::None || cancellation.IsCancellationRequested())
        {
            result.hr = DescribeStoppedWork(progress,
                                            result.errorText,
                                            &_cancelledOperationCount,
                                            &_workLimitFailureCount);
            return result;
        }
        result.hr = E_FAIL;
        const std::wstring context =
            FormatEngineString(IDS_VIEWERSQLITE_ERROR_PREPARE_PREVIEW_FMT, SanitizeIdentifierText(exactTable));
        result.errorText = DescribeSqliteFailure(_connection.get(), rc, context);
        return result;
    }

    result.hr = ReadQueryPage(_connection.get(),
                              stmt,
                              sql,
                              requestedPageSize,
                              rowOffset,
                              false,
                              budget,
                              progress,
                              result.page,
                              result.errorText,
                              &_cancelledOperationCount,
                              &_workLimitFailureCount);
    Debug::Perf::Emit(L"viewer.sqlite.page_us",
                      SUCCEEDED(result.hr) ? L"table" : (result.hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) ? L"table-cancelled" : L"table-failed"),
                      Debug::Perf::ElapsedUs(startedAt),
                      rowOffset,
                      static_cast<uint64_t>(result.page.rows.size()),
                      result.hr);
    return result;
}

QueryPageResult DatabaseSource::ExecuteReadOnlyQuery(std::wstring_view sql,
                                                     const uint32_t rowCap,
                                                     const QueryCancellation cancellation,
                                                     const QueryWorkBudget budget) const noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    QueryPageResult result{};
    result.hr = CheckCancellation(cancellation, result.errorText, &_cancelledOperationCount);
    if (FAILED(result.hr))
    {
        return result;
    }

    std::scoped_lock lock(_connectionMutex);
    ConnectionUseCounter connectionUse(_activeConnectionUse, _maxConcurrentConnectionUse);
    static_cast<void>(_operationCount.fetch_add(1u, std::memory_order_relaxed));
    if (! _connection)
    {
        result.hr        = E_UNEXPECTED;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_CONNECTION_UNAVAILABLE);
        return result;
    }

    result.hr = CheckCancellation(cancellation, result.errorText, &_cancelledOperationCount);
    if (FAILED(result.hr))
    {
        return result;
    }

    SqliteProgressState progress{.cancellation = cancellation, .budget = budget};
    ScopedSqliteProgress progressHandler(_connection.get(), progress);
    unique_sqlite_stmt stmt(nullptr, sqlite3_finalize);
    std::wstring executedSql;
    result.hr = PrepareSingleReadonlyStatement(_connection.get(),
                                               sql,
                                               stmt,
                                               executedSql,
                                               result.errorText,
                                               progress,
                                               &_cancelledOperationCount,
                                               &_workLimitFailureCount);
    if (FAILED(result.hr))
    {
        return result;
    }

    result.hr = ReadQueryPage(_connection.get(),
                              stmt,
                              std::move(executedSql),
                              std::clamp<uint32_t>(rowCap, 1u, kMaxQueryRowCap),
                              0,
                              true,
                              budget,
                              progress,
                              result.page,
                              result.errorText,
                              &_cancelledOperationCount,
                              &_workLimitFailureCount);
    Debug::Perf::Emit(L"viewer.sqlite.page_us",
                      SUCCEEDED(result.hr) ? L"query" : (result.hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) ? L"query-cancelled" : L"query-failed"),
                      Debug::Perf::ElapsedUs(startedAt),
                      0u,
                      static_cast<uint64_t>(result.page.rows.size()),
                      result.hr);
    return result;
}

ValidationResult DatabaseSource::ValidateReadOnlyQuery(std::wstring_view sql,
                                                       const QueryCancellation cancellation,
                                                       const QueryWorkBudget budget) const noexcept
{
    ValidationResult result{};
    result.hr = CheckCancellation(cancellation, result.errorText, &_cancelledOperationCount);
    if (FAILED(result.hr))
    {
        return result;
    }

    std::scoped_lock lock(_connectionMutex);
    ConnectionUseCounter connectionUse(_activeConnectionUse, _maxConcurrentConnectionUse);
    static_cast<void>(_operationCount.fetch_add(1u, std::memory_order_relaxed));
    if (! _connection)
    {
        result.hr        = E_UNEXPECTED;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_CONNECTION_UNAVAILABLE);
        return result;
    }

    result.hr = CheckCancellation(cancellation, result.errorText, &_cancelledOperationCount);
    if (FAILED(result.hr))
    {
        return result;
    }

    SqliteProgressState progress{.cancellation = cancellation, .budget = budget};
    ScopedSqliteProgress progressHandler(_connection.get(), progress);
    unique_sqlite_stmt stmt(nullptr, sqlite3_finalize);
    std::wstring executedSql;
    result.hr = PrepareSingleReadonlyStatement(_connection.get(),
                                               sql,
                                               stmt,
                                               executedSql,
                                               result.errorText,
                                               progress,
                                               &_cancelledOperationCount,
                                               &_workLimitFailureCount);
    if (SUCCEEDED(result.hr))
    {
        result.accepted = true;
    }
    return result;
}

SourceOpenResult DatabaseSource::OpenOwnedSnapshot(std::filesystem::path snapshotPath,
                                                   std::wstring displayName,
                                                   wil::unique_handle snapshotLifetimeHandle,
                                                   const SnapshotKind snapshotKind,
                                                   const uint64_t snapshotBytes,
                                                   const QueryCancellation cancellation) noexcept
{
    SourceOpenResult result{};
    unique_sqlite3 connection;
    result.hr = OpenReadOnlyConnection(snapshotPath, connection, result.errorText);
    if (FAILED(result.hr))
    {
        return result;
    }

    auto source = std::shared_ptr<DatabaseSource>(new (std::nothrow)
                                                      DatabaseSource(std::move(snapshotPath),
                                                                     std::move(displayName),
                                                                     std::move(snapshotLifetimeHandle),
                                                                     std::move(connection),
                                                                     snapshotKind,
                                                                     snapshotBytes));
    if (! source)
    {
        result.hr        = E_OUTOFMEMORY;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_SOURCE_ALLOCATE);
        return result;
    }

    result.hr = source->ListTables(result.tables, result.errorText, cancellation);
    if (FAILED(result.hr))
    {
        return result;
    }

    result.hr     = S_OK;
    result.source = std::move(source);
    return result;
}

SourceOpenResult DatabaseSource::OpenFromPath(std::filesystem::path localPath,
                                              std::wstring displayName,
                                              const QueryCancellation cancellation,
                                              const uint64_t maxSnapshotBytes) noexcept
{
    SourceOpenResult result{};
    TempSnapshot snapshot{};
    result.hr = CreateCrashCleanedTempSnapshot(snapshot, result.errorText);
    if (FAILED(result.hr))
    {
        return result;
    }

    uint64_t snapshotBytes = 0;
    result.hr = BackupLocalDatabase(localPath, snapshot, cancellation, maxSnapshotBytes, snapshotBytes, result.errorText);
    if (FAILED(result.hr))
    {
        return result;
    }

    return OpenOwnedSnapshot(std::move(snapshot.path),
                             std::move(displayName),
                             std::move(snapshot.lifetimeHandle),
                             SnapshotKind::LocalSqliteBackup,
                             snapshotBytes,
                             cancellation);
}

SourceOpenResult OpenFromViewerContext(IFileSystem* fileSystem,
                                       std::wstring_view path,
                                       const bool directOpenLocalFiles,
                                       const QueryCancellation cancellation) noexcept
{
    SourceOpenResult result{};
    if (fileSystem == nullptr || path.empty())
    {
        result.hr        = E_INVALIDARG;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_CONTEXT_MISSING);
        return result;
    }

    static_cast<void>(directOpenLocalFiles); // Legacy configuration: local databases are always isolated through SQLite backup.
    const std::filesystem::path candidate(path);
    if (candidate.is_absolute() && IsBuiltinFileSystem(fileSystem))
    {
        return DatabaseSource::OpenFromPath(candidate, candidate.filename().wstring(), cancellation);
    }

    wil::com_ptr<IFileSystemIO> fileIo;
    const HRESULT fileIoHr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (FAILED(fileIoHr) || ! fileIo)
    {
        result.hr        = FAILED(fileIoHr) ? fileIoHr : E_NOINTERFACE;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_FILESYSTEM_READ_UNSUPPORTED);
        return result;
    }

    wil::com_ptr<IFileReader> reader;
    const std::wstring pathText(path);
    const HRESULT readerHr = fileIo->CreateFileReader(pathText.c_str(), reader.put());
    if (FAILED(readerHr) || ! reader)
    {
        result.hr        = FAILED(readerHr) ? readerHr : E_FAIL;
        result.errorText = ReadEngineString(IDS_VIEWERSQLITE_ERROR_FILESYSTEM_OPEN);
        return result;
    }

    TempSnapshot snapshot{};
    result.hr = CreateCrashCleanedTempSnapshot(snapshot, result.errorText);
    if (FAILED(result.hr))
    {
        return result;
    }

    uint64_t snapshotBytes = 0;
    result.hr = CopyReaderToSnapshot(reader.get(), snapshot, cancellation, snapshotBytes, result.errorText);
    if (FAILED(result.hr))
    {
        return result;
    }

    const std::filesystem::path displayPath(pathText);
    return DatabaseSource::OpenOwnedSnapshot(std::move(snapshot.path),
                                             displayPath.filename().wstring(),
                                             std::move(snapshot.lifetimeHandle),
                                             SnapshotKind::VirtualByteCopy,
                                             snapshotBytes,
                                             cancellation);
}

std::wstring QuoteIdentifier(std::wstring_view name)
{
    std::wstring quoted;
    quoted.reserve(name.size() + 2u);
    quoted.push_back(L'"');
    for (const wchar_t ch : name)
    {
        quoted.push_back(ch);
        if (ch == L'"')
        {
            quoted.push_back(L'"');
        }
    }
    quoted.push_back(L'"');
    return quoted;
}

std::wstring BuildTablePreviewSql(std::wstring_view tableName,
                                  const uint32_t pageSize,
                                  const uint64_t rowOffset,
                                  const size_t orderByColumnIndex,
                                  const TableSortDirection sortDirection)
{
    std::wstring sql = std::format(L"SELECT * FROM {}", QuoteIdentifier(tableName));
    if (sortDirection != TableSortDirection::None && orderByColumnIndex != kNoSortColumn)
    {
        sql.append(std::format(L" ORDER BY {} {}", orderByColumnIndex + 1u, sortDirection == TableSortDirection::Ascending ? L"ASC" : L"DESC"));
    }
    sql.append(std::format(L" LIMIT {} OFFSET {}", pageSize, rowOffset));
    return sql;
}
} // namespace ViewerSqliteEngine
