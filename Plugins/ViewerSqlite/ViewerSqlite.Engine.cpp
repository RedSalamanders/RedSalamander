#include "ViewerSqlite.Engine.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <cwctype>
#include <format>
#include <limits>
#include <memory>
#include <new>
#include <string_view>
#include <system_error>

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <sqlite3.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

namespace ViewerSqliteEngine
{
namespace
{
using unique_sqlite3     = std::unique_ptr<sqlite3, decltype(&sqlite3_close_v2)>;
using unique_sqlite_stmt = std::unique_ptr<sqlite3_stmt, decltype(&sqlite3_finalize)>;

constexpr size_t kCopyChunkBytes   = 1024u * 1024u;
constexpr size_t kMaxCellChars     = 4096u;
constexpr uint32_t kMaxPageSize    = 1000u;
constexpr uint32_t kMaxQueryRowCap = 100000u;

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
    if (text.empty())
    {
        return {};
    }

    const int needed = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (needed <= 0)
    {
        return {};
    }

    std::wstring out(static_cast<size_t>(needed), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), out.data(), needed);
    if (written != needed)
    {
        return {};
    }

    return out;
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int needed = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (needed <= 0)
    {
        return {};
    }

    std::string out(static_cast<size_t>(needed), '\0');
    const int written = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), out.data(), needed, nullptr, nullptr);
    if (written != needed)
    {
        return {};
    }

    return out;
}

[[nodiscard]] std::wstring DescribeSqliteFailure(sqlite3* db, int sqliteCode, std::wstring_view context) noexcept
{
    const char* raw     = (db != nullptr) ? sqlite3_errmsg(db) : sqlite3_errstr(sqliteCode);
    std::wstring detail = Utf16FromUtf8(raw != nullptr ? std::string_view(raw) : std::string_view{});
    if (detail.empty())
    {
        detail = std::format(L"SQLite error {}", sqliteCode);
    }

    if (context.empty())
    {
        return detail;
    }

    return std::format(L"{}: {}", context, detail);
}

[[nodiscard]] HRESULT OpenReadOnlyConnection(const std::filesystem::path& localPath, unique_sqlite3& db, std::wstring& errorText) noexcept
{
    const std::string pathUtf8 = Utf8FromUtf16(localPath.wstring());
    if (pathUtf8.empty())
    {
        errorText = L"Failed to convert the database path to UTF-8.";
        return E_INVALIDARG;
    }

    sqlite3* raw = nullptr;
    const int rc = sqlite3_open_v2(pathUtf8.c_str(), &raw, SQLITE_OPEN_READONLY | SQLITE_OPEN_NOMUTEX | SQLITE_OPEN_PRIVATECACHE, nullptr);
    db.reset(raw);
    if (rc != SQLITE_OK || ! db)
    {
        errorText = DescribeSqliteFailure(db.get(), rc, L"Failed to open database");
        return E_FAIL;
    }

    static_cast<void>(sqlite3_extended_result_codes(db.get(), 1));
    return S_OK;
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

    if (value.size() > kMaxCellChars)
    {
        size_t truncatedLength = kMaxCellChars;
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

    return value;
}

[[nodiscard]] std::wstring ReadCellValue(sqlite3_stmt* stmt, int columnIndex) noexcept
{
    const int type = sqlite3_column_type(stmt, columnIndex);
    switch (type)
    {
        case SQLITE_NULL: return L"NULL";
        case SQLITE_BLOB:
        {
            const int bytes = sqlite3_column_bytes(stmt, columnIndex);
            return std::format(L"BLOB ({} bytes)", bytes);
        }
        case SQLITE_INTEGER:
        case SQLITE_FLOAT:
        {
            const unsigned char* text = sqlite3_column_text(stmt, columnIndex);
            if (text == nullptr)
            {
                return {};
            }

            return Utf16FromUtf8(std::string_view(reinterpret_cast<const char*>(text)));
        }
        case SQLITE_TEXT:
        default:
        {
            const void* text16 = sqlite3_column_text16(stmt, columnIndex);
            const int bytes16  = sqlite3_column_bytes16(stmt, columnIndex);
            if (text16 == nullptr || bytes16 <= 0)
            {
                return {};
            }

            const auto length = static_cast<size_t>(bytes16 / static_cast<int>(sizeof(wchar_t)));
            const auto* wide  = static_cast<const wchar_t*>(text16);
            return SanitizeCellText(std::wstring(wide, length));
        }
    }
}

[[nodiscard]] HRESULT ReadQueryPage(sqlite3* db,
                                    unique_sqlite_stmt& stmt,
                                    std::wstring executedSql,
                                    uint32_t rowLimit,
                                    uint64_t rowOffset,
                                    bool markTruncatedOnOverflow,
                                    QueryPage& page,
                                    std::wstring& errorText) noexcept
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
        columnInfo.name         = Utf16FromUtf8(rawName != nullptr ? std::string_view(rawName) : std::string_view{});
        columnInfo.declaredType = Utf16FromUtf8(rawDecl != nullptr ? std::string_view(rawDecl) : std::string_view{});
        page.columns.push_back(std::move(columnInfo));
    }

    const uint32_t effectiveLimit = std::clamp<uint32_t>(rowLimit, 1u, kMaxQueryRowCap);
    for (;;)
    {
        const int rc = sqlite3_step(stmt.get());
        if (rc == SQLITE_DONE)
        {
            break;
        }

        if (rc != SQLITE_ROW)
        {
            errorText = DescribeSqliteFailure(db, rc, L"Failed while reading query rows");
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
            row.push_back(ReadCellValue(stmt.get(), columnIndex));
        }

        page.rows.push_back(std::move(row));
    }

    return S_OK;
}

[[nodiscard]] HRESULT PrepareSingleReadonlyStatement(
    sqlite3* db, std::wstring_view sql, unique_sqlite_stmt& stmt, std::wstring& executedSql, std::wstring& errorText) noexcept
{
    executedSql = TrimWhitespace(sql);
    if (executedSql.empty())
    {
        errorText = L"Enter a SQL query.";
        return E_INVALIDARG;
    }

    const std::string sqlUtf8 = Utf8FromUtf16(executedSql);
    if (sqlUtf8.empty())
    {
        errorText = L"Failed to convert the SQL query to UTF-8.";
        return E_INVALIDARG;
    }

    sqlite3_stmt* raw = nullptr;
    const char* tail  = nullptr;
    const int rc      = sqlite3_prepare_v2(db, sqlUtf8.c_str(), static_cast<int>(sqlUtf8.size()), &raw, &tail);
    stmt.reset(raw);
    if (rc != SQLITE_OK)
    {
        errorText = DescribeSqliteFailure(db, rc, L"Failed to prepare query");
        return E_FAIL;
    }

    if (! stmt)
    {
        errorText = L"Enter a SQL query.";
        return E_INVALIDARG;
    }

    if (! TailHasOnlyWhitespaceOrSemicolons(tail))
    {
        errorText = L"Only a single SQL statement is allowed.";
        return E_INVALIDARG;
    }

    if (sqlite3_stmt_readonly(stmt.get()) == 0)
    {
        errorText = L"Only read-only SQL statements are allowed.";
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

[[nodiscard]] HRESULT CreateTempSnapshotPath(std::filesystem::path& pathOut, std::wstring& errorText) noexcept
{
    wchar_t tempDirBuffer[MAX_PATH + 1] = {};
    const DWORD tempDirLength           = GetTempPathW(static_cast<DWORD>(std::size(tempDirBuffer)), tempDirBuffer);
    if (tempDirLength == 0 || tempDirLength >= std::size(tempDirBuffer))
    {
        errorText = L"Failed to get the temporary directory.";
        return HRESULT_FROM_WIN32(GetLastError());
    }

    wchar_t tempFileBuffer[MAX_PATH + 1] = {};
    const UINT tempFileResult            = GetTempFileNameW(tempDirBuffer, L"rsq", 0, tempFileBuffer);
    if (tempFileResult == 0)
    {
        errorText = L"Failed to create a temporary file name.";
        return HRESULT_FROM_WIN32(GetLastError());
    }

    pathOut = std::filesystem::path(tempFileBuffer);
    return S_OK;
}

[[nodiscard]] HRESULT CopyReaderToSnapshot(IFileReader* reader, const std::filesystem::path& snapshotPath, std::wstring& errorText) noexcept
{
    if (reader == nullptr)
    {
        errorText = L"File reader is not available.";
        return E_INVALIDARG;
    }

    wil::unique_handle file(
        CreateFileW(snapshotPath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_SEQUENTIAL_SCAN, nullptr));
    if (! file)
    {
        errorText = L"Failed to create the local database snapshot.";
        return HRESULT_FROM_WIN32(GetLastError());
    }

    uint64_t ignored     = 0;
    const HRESULT seekHr = reader->Seek(0, FILE_BEGIN, &ignored);
    if (FAILED(seekHr))
    {
        errorText = L"Failed to seek to the beginning of the source database.";
        return seekHr;
    }

    auto buffer = std::unique_ptr<BYTE[]>(new (std::nothrow) BYTE[kCopyChunkBytes]);
    if (! buffer)
    {
        errorText = L"Out of memory while creating the local database snapshot.";
        return E_OUTOFMEMORY;
    }

    for (;;)
    {
        unsigned long readBytes = 0;
        const HRESULT readHr    = reader->Read(buffer.get(), static_cast<unsigned long>(kCopyChunkBytes), &readBytes);
        if (FAILED(readHr))
        {
            errorText = L"Failed to read the source database.";
            return readHr;
        }

        if (readBytes == 0)
        {
            break;
        }

        DWORD writtenBytes = 0;
        const BOOL writeOk = WriteFile(file.get(), buffer.get(), readBytes, &writtenBytes, nullptr);
        if (writeOk == 0 || writtenBytes != readBytes)
        {
            errorText = L"Failed to write the local database snapshot.";
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }

    return S_OK;
}

} // namespace

DatabaseSource::DatabaseSource(std::filesystem::path localPath, std::wstring displayName, bool deleteOnClose) noexcept
    : _localPath(std::move(localPath)),
      _displayName(std::move(displayName)),
      _deleteOnClose(deleteOnClose)
{
}

DatabaseSource::~DatabaseSource() noexcept
{
    if (! _deleteOnClose || _localPath.empty())
    {
        return;
    }

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

HRESULT DatabaseSource::ListTables(std::vector<TableInfo>& tables, std::wstring& errorText) const noexcept
{
    tables.clear();
    errorText.clear();

    unique_sqlite3 db(nullptr, sqlite3_close_v2);
    HRESULT hr = OpenReadOnlyConnection(_localPath, db, errorText);
    if (FAILED(hr))
    {
        return hr;
    }

    constexpr char kSql[] = "SELECT name, type "
                            "FROM sqlite_schema "
                            "WHERE type IN ('table', 'view') "
                            "  AND name NOT LIKE 'sqlite_%' "
                            "ORDER BY CASE type WHEN 'table' THEN 0 ELSE 1 END, name COLLATE NOCASE";

    sqlite3_stmt* raw = nullptr;
    const int rc      = sqlite3_prepare_v2(db.get(), kSql, -1, &raw, nullptr);
    unique_sqlite_stmt stmt(raw, sqlite3_finalize);
    if (rc != SQLITE_OK || ! stmt)
    {
        errorText = DescribeSqliteFailure(db.get(), rc, L"Failed to enumerate tables");
        return E_FAIL;
    }

    for (;;)
    {
        const int stepRc = sqlite3_step(stmt.get());
        if (stepRc == SQLITE_DONE)
        {
            break;
        }

        if (stepRc != SQLITE_ROW)
        {
            errorText = DescribeSqliteFailure(db.get(), stepRc, L"Failed to enumerate tables");
            return E_FAIL;
        }

        TableInfo table{};
        table.name = ReadCellValue(stmt.get(), 0);
        table.kind = ReadCellValue(stmt.get(), 1);
        if (! table.name.empty())
        {
            tables.push_back(std::move(table));
        }
    }

    return S_OK;
}

QueryPageResult DatabaseSource::LoadTablePage(
    std::wstring_view tableName, uint32_t pageSize, uint64_t rowOffset, size_t orderByColumnIndex, TableSortDirection sortDirection) const noexcept
{
    QueryPageResult result{};
    const std::wstring trimmedTable = TrimWhitespace(tableName);
    if (trimmedTable.empty())
    {
        result.hr        = E_INVALIDARG;
        result.errorText = L"Choose a table to preview.";
        return result;
    }

    unique_sqlite3 db(nullptr, sqlite3_close_v2);
    result.hr = OpenReadOnlyConnection(_localPath, db, result.errorText);
    if (FAILED(result.hr))
    {
        return result;
    }

    const uint32_t requestedPageSize = std::clamp<uint32_t>(pageSize, 1u, kMaxPageSize);
    const uint32_t fetchPageSize     = requestedPageSize + 1u;
    const std::wstring sql           = BuildTablePreviewSql(trimmedTable, fetchPageSize, rowOffset, orderByColumnIndex, sortDirection);
    const std::string sqlUtf8        = Utf8FromUtf16(sql);
    if (sqlUtf8.empty())
    {
        result.hr        = E_INVALIDARG;
        result.errorText = L"Failed to build the table preview SQL.";
        return result;
    }

    sqlite3_stmt* raw = nullptr;
    const int rc      = sqlite3_prepare_v2(db.get(), sqlUtf8.c_str(), static_cast<int>(sqlUtf8.size()), &raw, nullptr);
    unique_sqlite_stmt stmt(raw, sqlite3_finalize);
    if (rc != SQLITE_OK || ! stmt)
    {
        result.hr        = E_FAIL;
        result.errorText = DescribeSqliteFailure(db.get(), rc, std::format(L"Failed to prepare preview for '{}'", trimmedTable));
        return result;
    }

    result.hr = ReadQueryPage(db.get(), stmt, sql, requestedPageSize, rowOffset, false, result.page, result.errorText);
    return result;
}

QueryPageResult DatabaseSource::ExecuteReadOnlyQuery(std::wstring_view sql, uint32_t rowCap) const noexcept
{
    QueryPageResult result{};

    unique_sqlite3 db(nullptr, sqlite3_close_v2);
    result.hr = OpenReadOnlyConnection(_localPath, db, result.errorText);
    if (FAILED(result.hr))
    {
        return result;
    }

    unique_sqlite_stmt stmt(nullptr, sqlite3_finalize);
    std::wstring executedSql;
    result.hr = PrepareSingleReadonlyStatement(db.get(), sql, stmt, executedSql, result.errorText);
    if (FAILED(result.hr))
    {
        return result;
    }

    result.hr =
        ReadQueryPage(db.get(), stmt, std::move(executedSql), std::clamp<uint32_t>(rowCap, 1u, kMaxQueryRowCap), 0, true, result.page, result.errorText);
    return result;
}

ValidationResult DatabaseSource::ValidateReadOnlyQuery(std::wstring_view sql) const noexcept
{
    ValidationResult result{};

    unique_sqlite3 db(nullptr, sqlite3_close_v2);
    result.hr = OpenReadOnlyConnection(_localPath, db, result.errorText);
    if (FAILED(result.hr))
    {
        return result;
    }

    unique_sqlite_stmt stmt(nullptr, sqlite3_finalize);
    std::wstring executedSql;
    result.hr = PrepareSingleReadonlyStatement(db.get(), sql, stmt, executedSql, result.errorText);
    if (SUCCEEDED(result.hr))
    {
        result.accepted = true;
    }

    return result;
}

SourceOpenResult DatabaseSource::OpenFromPath(std::filesystem::path localPath, std::wstring displayName, bool deleteOnClose) noexcept
{
    SourceOpenResult result{};
    auto source = std::shared_ptr<DatabaseSource>(new (std::nothrow) DatabaseSource(std::move(localPath), std::move(displayName), deleteOnClose));
    if (! source)
    {
        result.hr        = E_OUTOFMEMORY;
        result.errorText = L"Failed to allocate the SQLite database source.";
        return result;
    }

    std::vector<TableInfo> tables;
    std::wstring errorText;
    const HRESULT listHr = source->ListTables(tables, errorText);
    if (FAILED(listHr))
    {
        result.hr        = listHr;
        result.errorText = std::move(errorText);
        return result;
    }

    result.hr     = S_OK;
    result.source = std::move(source);
    result.tables = std::move(tables);
    return result;
}

SourceOpenResult OpenFromViewerContext(IFileSystem* fileSystem, std::wstring_view path, bool directOpenLocalFiles) noexcept
{
    SourceOpenResult result{};
    if (fileSystem == nullptr || path.empty())
    {
        result.hr        = E_INVALIDARG;
        result.errorText = L"Viewer context is missing the file system or path.";
        return result;
    }

    if (directOpenLocalFiles)
    {
        const std::filesystem::path candidate(path);
        if (candidate.is_absolute() && IsBuiltinFileSystem(fileSystem))
        {
            return DatabaseSource::OpenFromPath(candidate, candidate.filename().wstring(), false);
        }
    }

    wil::com_ptr<IFileSystemIO> fileIo;
    const HRESULT fileIoHr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (FAILED(fileIoHr) || ! fileIo)
    {
        result.hr        = FAILED(fileIoHr) ? fileIoHr : E_NOINTERFACE;
        result.errorText = L"This file system does not support file reading for the SQLite viewer.";
        return result;
    }

    wil::com_ptr<IFileReader> reader;
    const std::wstring pathText(path);
    const HRESULT readerHr = fileIo->CreateFileReader(pathText.c_str(), reader.put());
    if (FAILED(readerHr) || ! reader)
    {
        result.hr        = FAILED(readerHr) ? readerHr : E_FAIL;
        result.errorText = L"Failed to open the database through the active file system.";
        return result;
    }

    std::filesystem::path snapshotPath;
    HRESULT snapshotHr = CreateTempSnapshotPath(snapshotPath, result.errorText);
    if (FAILED(snapshotHr))
    {
        result.hr = snapshotHr;
        return result;
    }

    snapshotHr = CopyReaderToSnapshot(reader.get(), snapshotPath, result.errorText);
    if (FAILED(snapshotHr))
    {
        std::error_code ec;
        std::filesystem::remove(snapshotPath, ec);
        result.hr = snapshotHr;
        return result;
    }

    const std::filesystem::path displayPath(pathText);
    return DatabaseSource::OpenFromPath(std::move(snapshotPath), displayPath.filename().wstring(), true);
}

std::wstring QuoteIdentifier(std::wstring_view name)
{
    std::wstring quoted;
    quoted.reserve(name.size() + 2u);
    quoted.push_back(L'"');
    for (wchar_t ch : name)
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

std::wstring BuildTablePreviewSql(
    std::wstring_view tableName, uint32_t pageSize, uint64_t rowOffset, size_t orderByColumnIndex, TableSortDirection sortDirection)
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
