#include "CrashHandler.h"

#include "Framework.h"

#include <algorithm>
#include <atomic>
#include <csignal>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <cwchar>
#include <exception>
#include <intrin.h>
#include <iterator>
#include <new>
#include <string_view>
#include <utility>
#include <vector>

#include <dbghelp.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include "AppDataPaths.h"

#pragma comment(lib, "Dbghelp.lib")

namespace CrashHandler
{
namespace
{
constexpr wchar_t kCompanyDirName[]          = L"RedSalamander";
constexpr wchar_t kCrashDirName[]            = L"Crashes";
constexpr wchar_t kMarkerFileName[]          = L"last_crash.txt";
constexpr std::wstring_view kOwnedDumpPrefix = L"RedSalamander-report-v1-";
constexpr size_t kKeepNewestReports          = 10;
constexpr size_t kDumpTimeDigits             = 20;
constexpr size_t kPathCapacity               = 32768;
constexpr size_t kCrashNameCapacity          = 96;

constexpr DWORD kTerminateCode        = 0xE000CAF1u;
constexpr DWORD kPureCallCode         = 0xE000CAF2u;
constexpr DWORD kInvalidParameterCode = 0xE000CAF3u;
constexpr DWORD kAbortCode            = 0xE000CAF4u;

std::atomic<bool> g_installed{false};
std::vector<wil::unique_hfile> g_crashDirectoryAnchors;
wil::unique_handle g_dumpFinished;
wil::unique_handle g_probePeer;
volatile LONG g_dumpThread = 0;
std::atomic<DWORD> g_exitCode{0};
wchar_t g_directory[kPathCapacity]{};
wchar_t g_markerPath[kPathCapacity]{};
wchar_t g_dumpPath[kPathCapacity]{};
wchar_t g_stagedMarkerPath[kPathCapacity]{};
alignas(FILE_RENAME_INFO) std::byte g_markerRename[offsetof(FILE_RENAME_INFO, FileName) + kPathCapacity * sizeof(wchar_t)]{};

[[nodiscard]] std::filesystem::path DefaultDirectory()
{
    const auto base = AppDataPaths::GetLocalAppDataPath();
    return base.empty() ? std::filesystem::path{} : base / kCompanyDirName / kCrashDirName;
}

[[nodiscard]] std::vector<wil::unique_hfile> PinLocalDirectory(const std::filesystem::path& directory, bool createProductDirectories)
{
    const auto rootName = directory.root_name().native();
    const auto root     = directory.root_path();
    if (! directory.is_absolute() || rootName.size() != 2 || rootName[1] != L':' || GetDriveTypeW(root.c_str()) != DRIVE_FIXED)
        return {};

    // Pin each component without sharing delete. The writer can then use fixed paths without following reparses.
    auto entered                = root;
    const auto productDirectory = directory.parent_path();
    std::vector<wil::unique_hfile> anchors;
    for (const auto& component : directory.relative_path())
    {
        if (component.empty() || component == L"." || component == L"..")
            return {};
        entered /= component;
        const auto open = [&]() noexcept
        {
            return CreateFileW(entered.c_str(),
                               FILE_LIST_DIRECTORY,
                               FILE_SHARE_READ | FILE_SHARE_WRITE,
                               nullptr,
                               OPEN_EXISTING,
                               FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                               nullptr);
        };
        wil::unique_hfile folder(open());
        const DWORD openError = folder ? ERROR_SUCCESS : GetLastError();
        if (createProductDirectories && (entered == productDirectory || entered == directory) &&
            (openError == ERROR_FILE_NOT_FOUND || openError == ERROR_PATH_NOT_FOUND))
        {
            if (! CreateDirectoryW(entered.c_str(), nullptr) && GetLastError() != ERROR_ALREADY_EXISTS)
                return {};
            folder.reset(open());
        }
        BY_HANDLE_FILE_INFORMATION info{};
        if (! folder || ! GetFileInformationByHandle(folder.get(), &info) ||
            (info.dwFileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != FILE_ATTRIBUTE_DIRECTORY)
            return {};
        anchors.push_back(std::move(folder));
    }
    return anchors;
}

[[nodiscard]] unsigned long long CrashTimestamp() noexcept
{
    FILETIME time{};
    GetSystemTimePreciseAsFileTime(&time);
    return (static_cast<unsigned long long>(time.dwHighDateTime) << 32u) | time.dwLowDateTime;
}

[[nodiscard]] bool FormatDumpPath(const wchar_t* directory, wchar_t* output, size_t capacity, unsigned long long timestamp) noexcept
{
    return swprintf_s(output, capacity, L"%ls\\%ls%020llu-p%lu.dmp", directory, kOwnedDumpPrefix.data(), timestamp, GetCurrentProcessId()) > 0;
}

[[nodiscard]] bool IsOwnedDumpName(std::wstring_view name) noexcept
{
    constexpr std::wstring_view extension = L".dmp";
    if (! name.starts_with(kOwnedDumpPrefix) || ! name.ends_with(extension))
        return false;

    name.remove_prefix(kOwnedDumpPrefix.size());
    name.remove_suffix(extension.size());
    constexpr size_t pidOffset = kDumpTimeDigits + 2;
    if (name.size() <= pidOffset || name.size() > pidOffset + 10 || name.substr(kDumpTimeDigits, 2) != L"-p")
        return false;

    const auto isDigit = [](wchar_t ch) noexcept { return ch >= L'0' && ch <= L'9'; };
    return std::ranges::all_of(name.substr(0, kDumpTimeDigits), isDigit) && std::ranges::all_of(name.substr(pidOffset), isDigit);
}

[[nodiscard]] bool DeletePlainReportFile(const std::filesystem::path& path) noexcept
{
    wil::unique_hfile file(
        CreateFileW(path.c_str(), FILE_READ_ATTRIBUTES | DELETE, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT, nullptr));
    BY_HANDLE_FILE_INFORMATION info{};
    if (! file || ! GetFileInformationByHandle(file.get(), &info) || (info.dwFileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0)
        return false;

    FILE_DISPOSITION_INFO disposition{.DeleteFile = TRUE};
    return SetFileInformationByHandle(file.get(), FileDispositionInfo, &disposition, sizeof(disposition)) != FALSE;
}

[[nodiscard]] size_t PruneOldReports(const std::filesystem::path& directory, std::wstring_view pendingDump = {})
{
    const auto anchors = PinLocalDirectory(directory, false);
    if (anchors.empty())
        return 0;
    const auto pattern = directory / L"RedSalamander-report-v1-*.dmp";
    WIN32_FIND_DATAW found{};
    wil::unique_hfind search(FindFirstFileW(pattern.c_str(), &found));
    if (! search)
        return 0;

    std::vector<std::wstring> dumps;
    do
    {
        if ((found.dwFileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) == 0 && IsOwnedDumpName(found.cFileName))
            dumps.emplace_back(found.cFileName);
    } while (FindNextFileW(search.get(), &found));
    if (GetLastError() != ERROR_NO_MORE_FILES)
        return 0;
    search.reset();
    std::ranges::sort(dumps, [](const std::wstring& lhs, const std::wstring& rhs) noexcept { return lhs > rhs; });

    size_t deleted = 0;
    for (size_t index = kKeepNewestReports; index < dumps.size(); ++index)
    {
        auto path = directory / dumps[index];
        if (path.native() != pendingDump && DeletePlainReportFile(path))
        {
            ++deleted;
            path.replace_extension(L".txt");
            static_cast<void>(DeletePlainReportFile(path));
        }
    }
    return deleted;
}

[[nodiscard]] bool DeleteOwnedFile(HANDLE file) noexcept
{
    FILE_DISPOSITION_INFO disposition{.DeleteFile = TRUE};
    return SetFileInformationByHandle(file, FileDispositionInfo, &disposition, sizeof(disposition)) != FALSE;
}

[[nodiscard]] bool WriteMarkerContents(HANDLE file, std::wstring_view dump) noexcept
{
    if (dump.size() >= kPathCapacity)
        return false;
    const wchar_t bom = 0xFEFF;
    const DWORD bytes = static_cast<DWORD>(dump.size() * sizeof(wchar_t));
    DWORD written     = 0;
    return WriteFile(file, &bom, sizeof(bom), &written, nullptr) && written == sizeof(bom) && WriteFile(file, dump.data(), bytes, &written, nullptr) &&
           written == bytes && FlushFileBuffers(file);
}

[[nodiscard]] bool ReplacePendingMarker(HANDLE file) noexcept
{
    const size_t nameBytes = wcslen(g_markerPath) * sizeof(wchar_t);
    auto* const rename     = reinterpret_cast<FILE_RENAME_INFO*>(g_markerRename);
    rename->Flags          = FILE_RENAME_FLAG_REPLACE_IF_EXISTS | FILE_RENAME_FLAG_POSIX_SEMANTICS;
    rename->RootDirectory  = nullptr;
    rename->FileNameLength = static_cast<DWORD>(nameBytes);
    std::memcpy(rename->FileName, g_markerPath, nameBytes + sizeof(wchar_t));
    const DWORD bytes = static_cast<DWORD>(offsetof(FILE_RENAME_INFO, FileName) + nameBytes + sizeof(wchar_t));
    if (SetFileInformationByHandle(file, FileRenameInfoEx, rename, bytes))
        return true;
    const DWORD error = GetLastError();
    if (error != ERROR_INVALID_PARAMETER && error != ERROR_INVALID_FUNCTION && error != ERROR_NOT_SUPPORTED && error != ERROR_CALL_NOT_IMPLEMENTED)
        return false;
    rename->Flags           = 0;
    rename->ReplaceIfExists = TRUE;
    return SetFileInformationByHandle(file, FileRenameInfo, rename, bytes) != FALSE;
}

// Caller owns the single fault-writer slot. DbgHelp is linked without delay loading.
void WriteDumpAndMarker(EXCEPTION_POINTERS* exception, unsigned long long timestamp) noexcept
{
    if (! FormatDumpPath(g_directory, g_dumpPath, std::size(g_dumpPath), timestamp))
        return;

    const HANDLE dump = CreateFileW(g_dumpPath, GENERIC_WRITE | DELETE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (dump == INVALID_HANDLE_VALUE)
        return;
    MINIDUMP_EXCEPTION_INFORMATION information{};
    information.ThreadId          = GetCurrentThreadId();
    information.ExceptionPointers = exception;
    information.ClientPointers    = FALSE;
    constexpr auto dumpType       = static_cast<MINIDUMP_TYPE>(MiniDumpWithIndirectlyReferencedMemory | MiniDumpScanMemory | MiniDumpWithThreadInfo |
                                                               MiniDumpWithUnloadedModules | MiniDumpWithHandleData);
    const bool completed =
        MiniDumpWriteDump(GetCurrentProcess(), GetCurrentProcessId(), dump, dumpType, exception ? &information : nullptr, nullptr, nullptr) &&
        FlushFileBuffers(dump);
    if (! completed)
        static_cast<void>(DeleteOwnedFile(dump));
    static_cast<void>(CloseHandle(dump));
    if (! completed ||
        swprintf_s(g_stagedMarkerPath, std::size(g_stagedMarkerPath), L"%ls\\pending-%lu-%020llu.tmp", g_directory, GetCurrentProcessId(), timestamp) <= 0)
        return;

    const HANDLE marker = CreateFileW(g_stagedMarkerPath, GENERIC_WRITE | DELETE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (marker == INVALID_HANDLE_VALUE)
        return;
    if (! WriteMarkerContents(marker, g_dumpPath) || ! ReplacePendingMarker(marker))
        static_cast<void>(DeleteOwnedFile(marker));
    static_cast<void>(CloseHandle(marker));
}

[[noreturn]] void ExitAfterCrash(DWORD code) noexcept
{
    static_cast<void>(TerminateProcess(GetCurrentProcess(), code));
    __fastfail(FAST_FAIL_FATAL_APP_EXIT);
}

[[noreturn]] void ReportFatal(EXCEPTION_POINTERS* exception, DWORD code) noexcept
{
    const auto thread = static_cast<LONG>(GetCurrentThreadId());
    const LONG writer = InterlockedCompareExchange(&g_dumpThread, thread, 0);
    if (writer == thread)
        ExitAfterCrash(code);
    if (writer != 0)
    {
        if (g_probePeer)
            static_cast<void>(SetEvent(g_probePeer.get()));
        static_cast<void>(WaitForSingleObject(g_dumpFinished.get(), INFINITE));
        ExitAfterCrash(g_exitCode.load(std::memory_order_acquire));
    }
    g_exitCode.store(code, std::memory_order_release);
    if (g_probePeer && WaitForSingleObject(g_probePeer.get(), 5000) != WAIT_OBJECT_0)
        ExitAfterCrash(ERROR_TIMEOUT);
    WriteDumpAndMarker(exception, CrashTimestamp());
    static_cast<void>(SetEvent(g_dumpFinished.get()));
    ExitAfterCrash(code);
}

[[noreturn]] void ReportCrtFailure(DWORD code) noexcept
{
    CONTEXT captured{};
    RtlCaptureContext(&captured);
    EXCEPTION_RECORD record{};
    record.ExceptionCode    = code;
    record.ExceptionFlags   = EXCEPTION_NONCONTINUABLE;
    record.ExceptionAddress = _ReturnAddress();
    EXCEPTION_POINTERS exception{&record, &captured};
    ReportFatal(&exception, code);
}

LONG WINAPI UnhandledFilter(EXCEPTION_POINTERS* exception) noexcept
{
    return WriteDumpForException(exception);
}

void OnTerminate() noexcept
{
    ReportCrtFailure(kTerminateCode);
}

void __cdecl OnPureCall() noexcept
{
    ReportCrtFailure(kPureCallCode);
}

void __cdecl OnInvalidParameter(const wchar_t*, const wchar_t*, const wchar_t*, unsigned int, uintptr_t) noexcept
{
    ReportCrtFailure(kInvalidParameterCode);
}

void __cdecl OnAbort(int) noexcept
{
    ReportCrtFailure(kAbortCode);
}

[[nodiscard]] HRESULT InstallHandlers(const std::filesystem::path& directory, bool createProductDirectories) noexcept
{
    try
    {
        if (g_installed.load(std::memory_order_acquire))
            return HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);
        if (directory.native().size() >= kPathCapacity - kCrashNameCapacity)
            return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
        auto anchors = PinLocalDirectory(directory, createProductDirectories);
        if (anchors.empty())
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        if (wcscpy_s(g_directory, directory.c_str()) != 0 || swprintf_s(g_markerPath, std::size(g_markerPath), L"%ls\\%ls", g_directory, kMarkerFileName) <= 0)
            return E_INVALIDARG;

        g_dumpFinished.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
        if (! g_dumpFinished)
            return HRESULT_FROM_WIN32(GetLastError());
        if (std::signal(SIGABRT, OnAbort) == SIG_ERR)
            return E_FAIL;
        static_cast<void>(_set_abort_behavior(0, _WRITE_ABORT_MSG | _CALL_REPORTFAULT));
        g_crashDirectoryAnchors = std::move(anchors);
        g_installed.store(true, std::memory_order_release);
        static_cast<void>(SetUnhandledExceptionFilter(UnhandledFilter));
        static_cast<void>(std::set_terminate(OnTerminate));
        static_cast<void>(_set_purecall_handler(OnPureCall));
        static_cast<void>(_set_invalid_parameter_handler(OnInvalidParameter));
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        return E_OUTOFMEMORY;
    }
    catch (const std::exception&)
    {
        return E_FAIL;
    }
}

[[nodiscard]] std::wstring ReadMarkerDumpPath(HANDLE marker)
{
    LARGE_INTEGER length{};
    if (! GetFileSizeEx(marker, &length) || length.QuadPart < static_cast<LONGLONG>(sizeof(wchar_t)) ||
        length.QuadPart > static_cast<LONGLONG>(kPathCapacity * sizeof(wchar_t)) || (length.QuadPart % sizeof(wchar_t)) != 0)
        return {};
    const auto bytes = static_cast<DWORD>(length.QuadPart);
    std::wstring content(bytes / sizeof(wchar_t), L'\0');
    DWORD read = 0;
    if (! ReadFile(marker, content.data(), bytes, &read, nullptr) || read != bytes)
        return {};
    if (! content.empty() && content.front() == 0xFEFF)
        content.erase(0, 1);
    while (! content.empty() &&
           (content.back() == L'\0' || content.back() == L'\r' || content.back() == L'\n' || content.back() == L' ' || content.back() == L'\t'))
        content.pop_back();
    return content;
}

[[nodiscard]] bool IsLocalDump(const std::filesystem::path& directory, std::wstring_view dump)
{
    if (dump.empty() || dump.find(L'\0') != std::wstring_view::npos)
        return false;
    const std::filesystem::path path(dump);
    if (path.parent_path() != directory || path.extension() != L".dmp")
        return false;
    const DWORD attributes = GetFileAttributesW(path.c_str());
    return attributes != INVALID_FILE_ATTRIBUTES && (attributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) == 0;
}

[[nodiscard]] std::optional<PendingReport> TakePendingFrom(const std::filesystem::path& directory)
{
    const auto anchors = PinLocalDirectory(directory, false);
    if (anchors.empty())
        return std::nullopt;

    const auto path = directory / kMarkerFileName;
    wil::unique_hfile marker(CreateFileW(path.c_str(), GENERIC_READ | DELETE, 0, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT, nullptr));
    if (! marker)
    {
        if (GetLastError() == ERROR_FILE_NOT_FOUND)
            static_cast<void>(PruneOldReports(directory));
        return std::nullopt;
    }
    BY_HANDLE_FILE_INFORMATION info{};
    if (! GetFileInformationByHandle(marker.get(), &info) || (info.dwFileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0)
        return std::nullopt;

    PendingReport report{.directory = directory, .dumpPath = ReadMarkerDumpPath(marker.get())};
    if (! IsLocalDump(directory, report.dumpPath))
        report.dumpPath.clear();
    if (! DeleteOwnedFile(marker.get()))
        return std::nullopt;
    marker.reset();

    try
    {
        static_cast<void>(PruneOldReports(directory, report.dumpPath));
    }
    catch (const std::exception&)
    {
        // The report has already been claimed. Retention can retry on the next normal startup.
    }
    return report;
}

#if defined(RS_ASAN_DEBUG_BUILD) || defined(__SANITIZE_ADDRESS__)
constexpr DWORD kProbeCode = 0xE000CAFEu;
#else
constexpr DWORD kProbeCode = EXCEPTION_ACCESS_VIOLATION;
#endif

[[noreturn]] void RaiseProbe() noexcept
{
    RaiseException(kProbeCode, EXCEPTION_NONCONTINUABLE, 0, nullptr);
    ExitAfterCrash(kProbeCode);
}

DWORD WINAPI ProbeWorker(void* start) noexcept
{
    if (WaitForSingleObject(static_cast<HANDLE>(start), INFINITE) != WAIT_OBJECT_0)
        return 1;
    __try
    {
        RaiseProbe();
    }
    __except (WriteDumpForException(GetExceptionInformation()))
    {
        ExitAfterCrash(ERROR_UNHANDLED_EXCEPTION);
    }
}
} // namespace

HRESULT Install() noexcept
{
    try
    {
        return InstallHandlers(DefaultDirectory(), true);
    }
    catch (const std::bad_alloc&)
    {
        return E_OUTOFMEMORY;
    }
    catch (const std::exception&)
    {
        return E_FAIL;
    }
}

HRESULT InstallForProbe(const std::filesystem::path& directory) noexcept
{
    return InstallHandlers(directory, false);
}

int WriteDumpForException(EXCEPTION_POINTERS* exception) noexcept
{
    if (g_installed.load(std::memory_order_acquire))
        ReportFatal(exception, exception && exception->ExceptionRecord ? exception->ExceptionRecord->ExceptionCode : kTerminateCode);
    return EXCEPTION_EXECUTE_HANDLER;
}

std::optional<PendingReport> TakePendingReport()
{
    return TakePendingFrom(DefaultDirectory());
}

[[noreturn]] void TriggerCrashTest(bool overlapWorkers) noexcept
{
    if (! overlapWorkers)
        RaiseProbe();

    g_probePeer.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    wil::unique_handle start(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! g_probePeer || ! start)
        ExitAfterCrash(2);
    wil::unique_handle first(CreateThread(nullptr, 0, ProbeWorker, start.get(), 0, nullptr));
    wil::unique_handle second(CreateThread(nullptr, 0, ProbeWorker, start.get(), 0, nullptr));
    if (! first || ! second || ! SetEvent(start.get()))
        ExitAfterCrash(2);
    static_cast<void>(WaitForSingleObject(first.get(), INFINITE));
    ExitAfterCrash(2);
}

#ifdef ENABLE_TESTS
std::filesystem::path GetCrashMarkerPathForTest()
{
    const auto directory = DefaultDirectory();
    return directory.empty() ? std::filesystem::path{} : directory / kMarkerFileName;
}

std::filesystem::path BuildDumpPathForTest(const std::filesystem::path& dir)
{
    wchar_t path[kPathCapacity]{};
    return FormatDumpPath(dir.c_str(), path, std::size(path), CrashTimestamp()) ? std::filesystem::path(path) : std::filesystem::path{};
}

size_t PruneOldReportsForTest(const std::filesystem::path& dir, std::wstring_view pendingDump)
{
    return PruneOldReports(dir, pendingDump);
}

HRESULT WriteMarkerFileForTest(const std::filesystem::path& markerPath, std::wstring_view dumpPath) noexcept
{
    if (markerPath.empty())
        return E_INVALIDARG;
    wil::unique_hfile file(CreateFileW(markerPath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
        return HRESULT_FROM_WIN32(GetLastError());
    return WriteMarkerContents(file.get(), dumpPath) ? S_OK : E_FAIL;
}

std::wstring ReadMarkerDumpPathForTest(const std::filesystem::path& markerPath)
{
    wil::unique_hfile marker(CreateFileW(markerPath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    return marker ? ReadMarkerDumpPath(marker.get()) : std::wstring{};
}

std::optional<PendingReport> TakePendingReportForTest(const std::filesystem::path& directory, bool interactive)
{
    return interactive ? TakePendingFrom(directory) : std::nullopt;
}
#endif
} // namespace CrashHandler
