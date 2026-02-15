#include "FolderWindow.FileOperationsInternal.h"

#include "FolderWindow.FileOperations.IssuesPane.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "SettingsSave.h"
#include "SettingsStore.h"

#include <bit>
#include <iterator>
#include <psapi.h>
#include <shellapi.h>
#include <thread>
#include <yyjson.h>

namespace
{
using Task = FolderWindow::FileOperationState::Task;

enum class ReparsePointPolicy : unsigned char
{
    CopyReparse,
    FollowTargets,
    Skip,
};

[[nodiscard]] ReparsePointPolicy ParseReparsePointPolicy(std::string_view text) noexcept
{
    if (text == "followTargets")
    {
        return ReparsePointPolicy::FollowTargets;
    }
    if (text == "skip")
    {
        return ReparsePointPolicy::Skip;
    }

    return ReparsePointPolicy::CopyReparse;
}

[[nodiscard]] ReparsePointPolicy GetReparsePointPolicyFromSettings(const Common::Settings::Settings& settings, const std::wstring& pluginId) noexcept
{
    const auto it = settings.plugins.configurationByPluginId.find(pluginId);
    if (it == settings.plugins.configurationByPluginId.end())
    {
        return ReparsePointPolicy::CopyReparse;
    }

    const Common::Settings::JsonValue& config = it->second;
    if (! std::holds_alternative<Common::Settings::JsonValue::ObjectPtr>(config.value))
    {
        return ReparsePointPolicy::CopyReparse;
    }

    const auto obj = std::get<Common::Settings::JsonValue::ObjectPtr>(config.value);
    if (! obj)
    {
        return ReparsePointPolicy::CopyReparse;
    }

    for (const auto& member : obj->members)
    {
        if (member.first != "reparsePointPolicy")
        {
            continue;
        }

        const Common::Settings::JsonValue& v = member.second;
        if (! std::holds_alternative<std::string>(v.value))
        {
            return ReparsePointPolicy::CopyReparse;
        }

        const std::string& text = std::get<std::string>(v.value);
        return ParseReparsePointPolicy(text);
    }

    return ReparsePointPolicy::CopyReparse;
}

constexpr std::wstring_view kFileOpsAppId                    = L"RedSalamander";
constexpr std::wstring_view kFileOpsIssuesPaneWindowId       = L"FileOperationsIssuesPane";
constexpr std::wstring_view kFileOpsPopupWindowId            = L"FileOperationsPopup";
constexpr std::wstring_view kDiagnosticsLogPrefix            = L"FileOperations-";
constexpr std::wstring_view kDiagnosticsLogExtension         = L".log";
constexpr std::wstring_view kDiagnosticsIssueReportPrefix    = L"FileOperations-Issues-";
constexpr std::wstring_view kDiagnosticsIssueReportExtension = L".txt";
constexpr size_t kMaxCompletedTaskSummaries                  = 24u;
constexpr size_t kMaxTaskIssueDiagnostics                    = 128u;
constexpr size_t kDefaultMaxDiagnosticsInMemory              = 256u;
constexpr size_t kDefaultMaxDiagnosticsPerFlush              = 64u;
constexpr size_t kDefaultMaxDiagnosticsLogFiles              = 14u;
constexpr size_t kDefaultMaxDiagnosticsIssueReportFiles      = 60u;
constexpr ULONGLONG kDefaultDiagnosticsFlushIntervalMs       = 5'000ull;
constexpr ULONGLONG kDefaultDiagnosticsCleanupIntervalMs     = 15ull * 60ull * 1000ull;

struct DiagnosticsSettings
{
    size_t maxDiagnosticsInMemory          = kDefaultMaxDiagnosticsInMemory;
    size_t maxDiagnosticsPerFlush          = kDefaultMaxDiagnosticsPerFlush;
    size_t maxDiagnosticsLogFiles          = kDefaultMaxDiagnosticsLogFiles;
    size_t maxDiagnosticsIssueReportFiles  = kDefaultMaxDiagnosticsIssueReportFiles;
    ULONGLONG diagnosticsFlushIntervalMs   = kDefaultDiagnosticsFlushIntervalMs;
    ULONGLONG diagnosticsCleanupIntervalMs = kDefaultDiagnosticsCleanupIntervalMs;
#if defined(_DEBUG) || defined(DEBUG)
    bool infoEnabled  = true;
    bool debugEnabled = true;
#else
    bool infoEnabled  = false;
    bool debugEnabled = false;
#endif
};

struct PreCalcProgressCookie
{
    std::mutex* totalsMutex          = nullptr;
    unsigned __int64* totalBytes     = nullptr;
    unsigned __int64* totalFiles     = nullptr;
    unsigned __int64* totalDirs      = nullptr;
    std::atomic<bool>* acceptUpdates = nullptr;
    unsigned __int64 lastBytes       = 0;
    unsigned __int64 lastFiles       = 0;
    unsigned __int64 lastDirs        = 0;
};

void UpdatePreCalcSnapshot(Task& task, unsigned __int64 totalBytes, unsigned __int64 totalFiles, unsigned __int64 totalDirs) noexcept
{
    constexpr unsigned __int64 maxUlong = static_cast<unsigned __int64>(std::numeric_limits<unsigned long>::max());
    task._preCalcTotalBytes.store(totalBytes, std::memory_order_release);
    task._preCalcFileCount.store(static_cast<unsigned long>(std::min(totalFiles, maxUlong)), std::memory_order_release);
    task._preCalcDirectoryCount.store(static_cast<unsigned long>(std::min(totalDirs, maxUlong)), std::memory_order_release);
}

[[nodiscard]] size_t GetPositiveSizeOrDefault(const std::optional<uint32_t>& value, size_t defaultValue) noexcept
{
    if (! value.has_value() || value.value() == 0)
    {
        return defaultValue;
    }

    return static_cast<size_t>(value.value());
}

[[nodiscard]] ULONGLONG GetPositiveIntervalOrDefault(const std::optional<uint32_t>& value, ULONGLONG defaultValue) noexcept
{
    if (! value.has_value() || value.value() == 0)
    {
        return defaultValue;
    }

    return static_cast<ULONGLONG>(value.value());
}

void CleanupDiagnosticsFilesInDirectory(const std::filesystem::path& directory,
                                        std::wstring_view filePrefix,
                                        std::wstring_view fileExtension,
                                        size_t maxFilesToKeep) noexcept
{
    if (directory.empty() || maxFilesToKeep == 0)
    {
        return;
    }

    std::error_code ec;
    std::vector<std::filesystem::path> files;
    for (std::filesystem::directory_iterator it(directory, ec), end; ! ec && it != end; it.increment(ec))
    {
        const std::filesystem::directory_entry& de = *it;
        if (! de.is_regular_file(ec))
        {
            continue;
        }

        const std::wstring fileNameText = de.path().filename().wstring();
        if (fileNameText.size() < (filePrefix.size() + fileExtension.size()))
        {
            continue;
        }
        if (fileNameText.rfind(filePrefix.data(), 0) != 0)
        {
            continue;
        }
        if (de.path().extension().wstring() != fileExtension)
        {
            continue;
        }

        files.push_back(de.path());
    }

    if (files.size() <= maxFilesToKeep)
    {
        return;
    }

    std::sort(files.begin(), files.end(), std::greater<std::filesystem::path>());
    for (size_t i = maxFilesToKeep; i < files.size(); ++i)
    {
        std::filesystem::remove(files[i], ec);
    }
}

[[nodiscard]] bool GetAutoDismissSuccessFromSettings(const Common::Settings::Settings& settings) noexcept
{
    if (! settings.fileOperations.has_value())
    {
        return false;
    }

    return settings.fileOperations->autoDismissSuccess;
}

void SetAutoDismissSuccessInSettings(Common::Settings::Settings& settings, bool enabled) noexcept
{
    if (settings.fileOperations.has_value())
    {
        settings.fileOperations->autoDismissSuccess = enabled;
    }
    else if (enabled)
    {
        settings.fileOperations.emplace();
        settings.fileOperations->autoDismissSuccess = true;
    }

    if (! settings.fileOperations.has_value())
    {
        return;
    }

    const Common::Settings::FileOperationsSettings defaults{};
    const auto& fileOperations = settings.fileOperations.value();
    const bool hasNonDefault   = fileOperations.autoDismissSuccess != defaults.autoDismissSuccess ||
                               fileOperations.maxDiagnosticsLogFiles != defaults.maxDiagnosticsLogFiles ||
                               fileOperations.diagnosticsInfoEnabled != defaults.diagnosticsInfoEnabled ||
                               fileOperations.diagnosticsDebugEnabled != defaults.diagnosticsDebugEnabled || fileOperations.maxIssueReportFiles.has_value() ||
                               fileOperations.maxDiagnosticsInMemory.has_value() || fileOperations.maxDiagnosticsPerFlush.has_value() ||
                               fileOperations.diagnosticsFlushIntervalMs.has_value() || fileOperations.diagnosticsCleanupIntervalMs.has_value();
    if (! hasNonDefault)
    {
        settings.fileOperations.reset();
    }
}

[[nodiscard]] DiagnosticsSettings GetDiagnosticsSettingsFromSettings(const Common::Settings::Settings* settings) noexcept
{
    DiagnosticsSettings diagnostics{};
    if (! settings || ! settings->fileOperations.has_value())
    {
        return diagnostics;
    }

    const auto& fileOperations                 = settings->fileOperations.value();
    diagnostics.maxDiagnosticsInMemory         = GetPositiveSizeOrDefault(fileOperations.maxDiagnosticsInMemory, diagnostics.maxDiagnosticsInMemory);
    diagnostics.maxDiagnosticsPerFlush         = GetPositiveSizeOrDefault(fileOperations.maxDiagnosticsPerFlush, diagnostics.maxDiagnosticsPerFlush);
    diagnostics.maxDiagnosticsLogFiles         = std::max<size_t>(1u, static_cast<size_t>(fileOperations.maxDiagnosticsLogFiles));
    diagnostics.maxDiagnosticsIssueReportFiles = GetPositiveSizeOrDefault(fileOperations.maxIssueReportFiles, diagnostics.maxDiagnosticsIssueReportFiles);
    diagnostics.diagnosticsFlushIntervalMs = GetPositiveIntervalOrDefault(fileOperations.diagnosticsFlushIntervalMs, diagnostics.diagnosticsFlushIntervalMs);
    diagnostics.diagnosticsCleanupIntervalMs =
        GetPositiveIntervalOrDefault(fileOperations.diagnosticsCleanupIntervalMs, diagnostics.diagnosticsCleanupIntervalMs);
    diagnostics.infoEnabled  = fileOperations.diagnosticsInfoEnabled;
    diagnostics.debugEnabled = fileOperations.diagnosticsDebugEnabled;
    return diagnostics;
}

[[nodiscard]] const wchar_t* OperationToString(FileSystemOperation operation) noexcept
{
    switch (operation)
    {
        case FILESYSTEM_COPY: return L"copy";
        case FILESYSTEM_MOVE: return L"move";
        case FILESYSTEM_DELETE: return L"delete";
        case FILESYSTEM_RENAME: return L"rename";
        default: return L"unknown";
    }
}

[[nodiscard]] bool IsCancellationStatus(HRESULT hr) noexcept
{
    return hr == E_ABORT || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

[[nodiscard]] const wchar_t* DiagnosticSeverityToString(FolderWindow::FileOperationState::DiagnosticSeverity severity) noexcept
{
    switch (severity)
    {
        case FolderWindow::FileOperationState::DiagnosticSeverity::Debug: return L"debug";
        case FolderWindow::FileOperationState::DiagnosticSeverity::Info: return L"info";
        case FolderWindow::FileOperationState::DiagnosticSeverity::Warning: return L"warning";
        case FolderWindow::FileOperationState::DiagnosticSeverity::Error: return L"error";
        default: return L"unknown";
    }
}

struct ProcessMemorySnapshot
{
    unsigned __int64 workingSetBytes = 0;
    unsigned __int64 privateBytes    = 0;
};

[[nodiscard]] ProcessMemorySnapshot CaptureProcessMemorySnapshot() noexcept
{
    ProcessMemorySnapshot snapshot{};

    PROCESS_MEMORY_COUNTERS_EX counters{};
    counters.cb = sizeof(counters);
    if (K32GetProcessMemoryInfo(GetCurrentProcess(), reinterpret_cast<PROCESS_MEMORY_COUNTERS*>(&counters), static_cast<DWORD>(sizeof(counters))) == 0)
    {
        return snapshot;
    }

    snapshot.workingSetBytes = static_cast<unsigned __int64>(counters.WorkingSetSize);
    snapshot.privateBytes    = static_cast<unsigned __int64>(counters.PrivateUsage);
    return snapshot;
}

[[nodiscard]] const wchar_t* Win32ErrorToSymbolicName(DWORD error) noexcept
{
    switch (error)
    {
        case ERROR_SUCCESS: return L"ERROR_SUCCESS";
        case ERROR_ACCESS_DENIED: return L"ERROR_ACCESS_DENIED";
        case ERROR_ALREADY_EXISTS: return L"ERROR_ALREADY_EXISTS";
        case ERROR_FILE_EXISTS: return L"ERROR_FILE_EXISTS";
        case ERROR_FILE_NOT_FOUND: return L"ERROR_FILE_NOT_FOUND";
        case ERROR_PATH_NOT_FOUND: return L"ERROR_PATH_NOT_FOUND";
        case ERROR_SHARING_VIOLATION: return L"ERROR_SHARING_VIOLATION";
        case ERROR_LOCK_VIOLATION: return L"ERROR_LOCK_VIOLATION";
        case ERROR_DISK_FULL: return L"ERROR_DISK_FULL";
        case ERROR_HANDLE_DISK_FULL: return L"ERROR_HANDLE_DISK_FULL";
        case ERROR_CANCELLED: return L"ERROR_CANCELLED";
        case ERROR_NOT_SUPPORTED: return L"ERROR_NOT_SUPPORTED";
        case ERROR_INVALID_NAME: return L"ERROR_INVALID_NAME";
        case ERROR_INVALID_PARAMETER: return L"ERROR_INVALID_PARAMETER";
        case ERROR_DIRECTORY: return L"ERROR_DIRECTORY";
        case ERROR_PARTIAL_COPY: return L"ERROR_PARTIAL_COPY";
        case ERROR_BAD_LENGTH: return L"ERROR_BAD_LENGTH";
        case ERROR_ARITHMETIC_OVERFLOW: return L"ERROR_ARITHMETIC_OVERFLOW";
        default: return nullptr;
    }
}

[[nodiscard]] std::wstring FormatDiagnosticHresultName(HRESULT hr) noexcept
{
    const wchar_t* known = nullptr;
    switch (hr)
    {
        case S_OK: known = L"S_OK"; break;
        case S_FALSE: known = L"S_FALSE"; break;
        case E_ABORT: known = L"E_ABORT"; break;
        case E_ACCESSDENIED: known = L"E_ACCESSDENIED"; break;
        case E_FAIL: known = L"E_FAIL"; break;
        case E_INVALIDARG: known = L"E_INVALIDARG"; break;
        case E_NOINTERFACE: known = L"E_NOINTERFACE"; break;
        case E_NOTIMPL: known = L"E_NOTIMPL"; break;
        case E_OUTOFMEMORY: known = L"E_OUTOFMEMORY"; break;
        case E_POINTER: known = L"E_POINTER"; break;
        case E_UNEXPECTED: known = L"E_UNEXPECTED"; break;
        default: break;
    }
    if (known)
    {
        return known;
    }

    if (HRESULT_FACILITY(hr) == FACILITY_WIN32)
    {
        const DWORD code = HRESULT_CODE(static_cast<DWORD>(hr));
        if (const wchar_t* win32Name = Win32ErrorToSymbolicName(code))
        {
            return win32Name;
        }

        return std::format(L"WIN32_ERROR_{}", static_cast<unsigned long>(code));
    }

    return std::format(L"HRESULT_0x{:08X}", static_cast<unsigned long>(hr));
}

[[nodiscard]] std::wstring FormatDiagnosticStatusText(HRESULT hr) noexcept
{
    wchar_t buffer[512]{};
    DWORD messageId = std::bit_cast<DWORD>(hr);
    if (HRESULT_FACILITY(hr) == FACILITY_WIN32)
    {
        const DWORD code = HRESULT_CODE(static_cast<DWORD>(hr));
        if (code != 0)
        {
            messageId = code;
        }
    }
    const DWORD written = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                                         nullptr,
                                         messageId,
                                         MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                                         buffer,
                                         static_cast<DWORD>(std::size(buffer)),
                                         nullptr);
    if (written == 0)
    {
        return std::format(L"HRESULT 0x{:08X}", static_cast<unsigned long>(hr));
    }

    std::wstring result(buffer, written);
    while (! result.empty())
    {
        const wchar_t ch = result.back();
        if (ch != L'\r' && ch != L'\n' && ch != L' ' && ch != L'\t')
        {
            break;
        }
        result.pop_back();
    }

    if (result.empty())
    {
        return std::format(L"HRESULT 0x{:08X}", static_cast<unsigned long>(hr));
    }

    return result;
}

[[nodiscard]] std::wstring EscapeDiagnosticField(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    std::wstring escaped;
    escaped.reserve(text.size());
    for (wchar_t ch : text)
    {
        if (ch == L'\r' || ch == L'\n' || ch == L'\t')
        {
            escaped.push_back(L' ');
        }
        else
        {
            escaped.push_back(ch);
        }
    }

    return escaped;
}

[[nodiscard]] std::wstring EscapeDiagnosticJsonString(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    std::wstring escaped;
    escaped.reserve(text.size());
    for (wchar_t ch : text)
    {
        switch (ch)
        {
            case L'\\': escaped.append(L"\\\\"); break;
            case L'"': escaped.append(L"\\\""); break;
            case L'\b': escaped.append(L"\\b"); break;
            case L'\f': escaped.append(L"\\f"); break;
            case L'\n': escaped.append(L"\\n"); break;
            case L'\r': escaped.append(L"\\r"); break;
            case L'\t': escaped.append(L"\\t"); break;
            default:
                if (ch < 0x20)
                {
                    std::format_to(std::back_inserter(escaped), L"\\u{:04X}", static_cast<unsigned>(ch));
                }
                else
                {
                    escaped.push_back(ch);
                }
                break;
        }
    }

    return escaped;
}

[[nodiscard]] std::wstring_view TrimTrailingSeparators(std::wstring_view path) noexcept
{
    while (! path.empty())
    {
        const wchar_t last = path.back();
        if (last != L'\\' && last != L'/')
        {
            break;
        }
        path.remove_suffix(1);
    }
    return path;
}

[[nodiscard]] bool IsSameOrChildPath(std::wstring_view root, std::wstring_view candidate) noexcept
{
    root      = TrimTrailingSeparators(root);
    candidate = TrimTrailingSeparators(candidate);

    if (root.empty() || candidate.size() < root.size())
    {
        return false;
    }

    if (root.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }

    const int prefixChars = static_cast<int>(root.size());
    if (CompareStringOrdinal(candidate.data(), prefixChars, root.data(), prefixChars, TRUE) != CSTR_EQUAL)
    {
        return false;
    }

    if (candidate.size() == root.size())
    {
        return true;
    }

    const wchar_t next = candidate[root.size()];
    return next == L'\\' || next == L'/';
}

[[nodiscard]] std::wstring_view GetPathLeaf(std::wstring_view path) noexcept
{
    const std::wstring_view trimmed = TrimTrailingSeparators(path);
    if (trimmed.empty())
    {
        return trimmed;
    }

    const size_t pos = trimmed.find_last_of(L"\\/");
    if (pos == std::wstring_view::npos)
    {
        return trimmed;
    }

    return trimmed.substr(pos + 1);
}

[[nodiscard]] wchar_t GuessPreferredSeparator(std::wstring_view folder) noexcept
{
    const bool hasForward = folder.find(L'/') != std::wstring_view::npos;
    const bool hasBack    = folder.find(L'\\') != std::wstring_view::npos;
    if (hasForward && ! hasBack)
    {
        return L'/';
    }
    return L'\\';
}

[[nodiscard]] std::wstring JoinFolderAndLeaf(std::wstring_view folder, std::wstring_view leaf) noexcept
{
    if (folder.empty())
    {
        return std::wstring(leaf);
    }

    std::wstring result(folder);
    const wchar_t sep = GuessPreferredSeparator(folder);
    if (! result.empty())
    {
        const wchar_t last = result.back();
        if (last != L'\\' && last != L'/')
        {
            result.push_back(sep);
        }
    }
    result.append(leaf);
    return result;
}

[[nodiscard]] unsigned int
DeterminePerItemMaxConcurrency(const wil::com_ptr<IFileSystem>& fileSystem, FileSystemOperation operation, FileSystemFlags flags, unsigned int uiMax) noexcept
{
    if (! fileSystem || uiMax == 0u)
    {
        return 1u;
    }

    const bool isCopyMove = operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE;
    const bool isDelete   = operation == FILESYSTEM_DELETE;
    if (! isCopyMove && ! isDelete)
    {
        return 1u;
    }

    const char* capabilitiesText = nullptr;
    if (FAILED(fileSystem->GetCapabilities(&capabilitiesText)) || ! capabilitiesText || capabilitiesText[0] == '\0')
    {
        return 1u;
    }

    const std::string_view capabilitiesView(capabilitiesText);
    std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> doc(
        yyjson_read(capabilitiesView.data(), capabilitiesView.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM), &yyjson_doc_free);
    if (! doc)
    {
        return 1u;
    }

    yyjson_val* root = yyjson_doc_get_root(doc.get());
    if (! root || ! yyjson_is_obj(root))
    {
        return 1u;
    }

    yyjson_val* concurrencyObject = yyjson_obj_get(root, "concurrency");
    if (! concurrencyObject || ! yyjson_is_obj(concurrencyObject))
    {
        return 1u;
    }

    const char* key = nullptr;
    if (isCopyMove)
    {
        key = "copyMoveMax";
    }
    else if (isDelete)
    {
        key = (flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0 ? "deleteRecycleBinMax" : "deleteMax";
    }

    if (! key)
    {
        return 1u;
    }

    yyjson_val* valueNode = yyjson_obj_get(concurrencyObject, key);
    if (! valueNode)
    {
        return 1u;
    }

    uint64_t concurrency = 0;
    if (yyjson_is_uint(valueNode))
    {
        concurrency = yyjson_get_uint(valueNode);
    }
    else if (yyjson_is_int(valueNode))
    {
        const int64_t signedValue = yyjson_get_int(valueNode);
        if (signedValue > 0)
        {
            concurrency = static_cast<uint64_t>(signedValue);
        }
    }

    if (concurrency == 0)
    {
        return 1u;
    }

    return std::clamp(static_cast<unsigned int>(std::min<uint64_t>(concurrency, static_cast<uint64_t>(uiMax))), 1u, uiMax);
}

[[nodiscard]] std::optional<DWORD> Win32ErrorFromHRESULT(HRESULT hr) noexcept
{
    if (hr == E_ACCESSDENIED)
    {
        return ERROR_ACCESS_DENIED;
    }
    if (hr == E_ABORT)
    {
        return ERROR_CANCELLED;
    }

    if (HRESULT_FACILITY(hr) == FACILITY_WIN32)
    {
        return HRESULT_CODE(hr);
    }

    return std::nullopt;
}

[[nodiscard]] bool IsNetworkOfflineError(DWORD error) noexcept
{
    switch (error)
    {
        case ERROR_BAD_NETPATH:
        case ERROR_BAD_NET_NAME:
        case ERROR_NETNAME_DELETED:
        case ERROR_NETWORK_UNREACHABLE:
        case ERROR_HOST_UNREACHABLE:
        case ERROR_PORT_UNREACHABLE:
        case ERROR_CONNECTION_UNAVAIL:
        case ERROR_NOT_CONNECTED:
        case ERROR_CONNECTION_REFUSED:
        case ERROR_NO_NETWORK:
        case ERROR_NETWORK_ACCESS_DENIED: return true;
        default: return false;
    }
}

[[nodiscard]] bool IsPathTooLongError(DWORD error) noexcept
{
    switch (error)
    {
        case ERROR_FILENAME_EXCED_RANGE:
        case ERROR_BUFFER_OVERFLOW: return true;
        default: return false;
    }
}

[[nodiscard]] bool IsCopyMoveOperation(FileSystemOperation operation) noexcept
{
    return operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE;
}

[[nodiscard]] bool IsDirectoryReparsePoint(const wil::com_ptr<IFileSystemIO>& fileSystemIo, std::wstring_view path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    unsigned long attributes = 0;
    if (fileSystemIo)
    {
        const HRESULT hr = fileSystemIo->GetAttributes(std::wstring(path).c_str(), &attributes);
        if (FAILED(hr))
        {
            return false;
        }
    }
    else
    {
        const DWORD win32 = GetFileAttributesW(std::wstring(path).c_str());
        if (win32 == INVALID_FILE_ATTRIBUTES)
        {
            return false;
        }
        attributes = win32;
    }

    return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
}

[[nodiscard]] Task::ConflictBucket ClassifyConflictBucket(FileSystemOperation operation,
                                                          FileSystemFlags flags,
                                                          const wil::com_ptr<IFileSystemIO>& fileSystemIo,
                                                          HRESULT status,
                                                          std::wstring_view sourcePath,
                                                          std::wstring_view destinationPath,
                                                          bool unsupportedReparseHint) noexcept
{
    if (status == HRESULT_FROM_WIN32(ERROR_CANCELLED) || status == E_ABORT)
    {
        return Task::ConflictBucket::Unknown;
    }

    if (unsupportedReparseHint)
    {
        return Task::ConflictBucket::UnsupportedReparse;
    }

    if (operation == FILESYSTEM_DELETE && (flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0)
    {
        // Deleting via the recycle bin is handled by the shell and can fail for a variety of reasons
        // (including cases that would succeed as a direct delete). Offer a permanent-delete fallback.
        return Task::ConflictBucket::RecycleBinFailed;
    }

    const std::optional<DWORD> errorOpt = Win32ErrorFromHRESULT(status);
    const DWORD error                   = errorOpt.value_or(0);

    switch (error)
    {
        case ERROR_ALREADY_EXISTS:
        case ERROR_FILE_EXISTS: return Task::ConflictBucket::Exists;
        case ERROR_SHARING_VIOLATION:
        case ERROR_LOCK_VIOLATION: return Task::ConflictBucket::SharingViolation;
        case ERROR_DISK_FULL:
        case ERROR_HANDLE_DISK_FULL: return Task::ConflictBucket::DiskFull;
        default: break;
    }

    if (IsPathTooLongError(error))
    {
        return Task::ConflictBucket::PathTooLong;
    }

    if (IsNetworkOfflineError(error))
    {
        return Task::ConflictBucket::NetworkOffline;
    }

    if (error == ERROR_NOT_SUPPORTED && IsCopyMoveOperation(operation) && IsDirectoryReparsePoint(fileSystemIo, sourcePath))
    {
        return Task::ConflictBucket::UnsupportedReparse;
    }

    if (error == ERROR_ACCESS_DENIED)
    {
        const bool isDelete           = operation == FILESYSTEM_DELETE;
        const std::wstring_view probe = isDelete ? sourcePath : destinationPath;

        if (! probe.empty())
        {
            unsigned long attributes = 0;
            bool gotAttributes       = false;
            if (fileSystemIo)
            {
                gotAttributes = SUCCEEDED(fileSystemIo->GetAttributes(std::wstring(probe).c_str(), &attributes));
            }
            else
            {
                const DWORD win32 = GetFileAttributesW(std::wstring(probe).c_str());
                if (win32 != INVALID_FILE_ATTRIBUTES)
                {
                    attributes    = win32;
                    gotAttributes = true;
                }
            }

            if (gotAttributes && (attributes & FILE_ATTRIBUTE_READONLY) != 0)
            {
                return Task::ConflictBucket::ReadOnly;
            }
        }

        return Task::ConflictBucket::AccessDenied;
    }

    return Task::ConflictBucket::Unknown;
}
} // namespace

FolderWindow::FileOperationState::Task::Task(FileOperationState& state) noexcept : _state(&state), _folderWindow(&state._owner)
{
    _conflictDecisionEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemProgress(FileSystemOperation operationType,
                                                                                     unsigned long totalItems,
                                                                                     unsigned long completedItems,
                                                                                     unsigned __int64 totalBytes,
                                                                                     unsigned __int64 completedBytes,
                                                                                     const wchar_t* currentSourcePath,
                                                                                     const wchar_t* currentDestinationPath,
                                                                                     unsigned __int64 currentItemTotalBytes,
                                                                                     unsigned __int64 currentItemCompletedBytes,
                                                                                     FileSystemOptions* options,
                                                                                     unsigned __int64 progressStreamId,
                                                                                     void* cookie) noexcept
{
    if (operationType != _operation)
    {
        return S_OK;
    }

    const ULONGLONG nowTick = GetTickCount64();

    {
        std::scoped_lock lock(_progressMutex);
        ++_progressCallbackCount;
        if (_executionMode == ExecutionMode::PerItem)
        {
            if (_perItemTotalItems > 0 && _operation != FILESYSTEM_DELETE)
            {
                _progressTotalItems = (std::max)(_progressTotalItems, _perItemTotalItems);
            }

            if (cookie != nullptr)
            {
                size_t found = _perItemInFlightCallCount;
                for (size_t i = 0; i < _perItemInFlightCallCount; ++i)
                {
                    if (_perItemInFlightCalls[i].cookie == cookie)
                    {
                        found = i;
                        break;
                    }
                }

                if (found < _perItemInFlightCallCount)
                {
                    _perItemInFlightCalls[found].completedItems = completedItems;
                    _perItemInFlightCalls[found].completedBytes = completedBytes;
                    if (totalItems > 0)
                    {
                        _perItemInFlightCalls[found].totalItems = (std::max)(_perItemInFlightCalls[found].totalItems, totalItems);
                    }
                }
                else
                {
                    if (_perItemInFlightCallCount < _perItemInFlightCalls.size())
                    {
                        _perItemInFlightCalls[_perItemInFlightCallCount].cookie         = cookie;
                        _perItemInFlightCalls[_perItemInFlightCallCount].completedItems = completedItems;
                        _perItemInFlightCalls[_perItemInFlightCallCount].completedBytes = completedBytes;
                        _perItemInFlightCalls[_perItemInFlightCallCount].totalItems     = totalItems;
                        ++_perItemInFlightCallCount;
                    }
                    else if (! _perItemInFlightCalls.empty())
                    {
                        _perItemInFlightCalls.back().cookie         = cookie;
                        _perItemInFlightCalls.back().completedItems = completedItems;
                        _perItemInFlightCalls.back().completedBytes = completedBytes;
                        _perItemInFlightCalls.back().totalItems     = totalItems;
                    }
                }
            }

            unsigned __int64 inFlightCompletedBytes = 0;
            unsigned __int64 inFlightCompletedItems = 0;
            unsigned __int64 inFlightTotalItems     = 0;
            for (size_t i = 0; i < _perItemInFlightCallCount; ++i)
            {
                const unsigned __int64 bytes = _perItemInFlightCalls[i].completedBytes;
                if (std::numeric_limits<unsigned __int64>::max() - inFlightCompletedBytes < bytes)
                {
                    inFlightCompletedBytes = std::numeric_limits<unsigned __int64>::max();
                    break;
                }
                inFlightCompletedBytes += bytes;

                const unsigned __int64 items = _perItemInFlightCalls[i].completedItems;
                if (std::numeric_limits<unsigned __int64>::max() - inFlightCompletedItems < items)
                {
                    inFlightCompletedItems = std::numeric_limits<unsigned __int64>::max();
                }
                else
                {
                    inFlightCompletedItems += items;
                }

                const unsigned __int64 total = static_cast<unsigned __int64>(_perItemInFlightCalls[i].totalItems);
                if (std::numeric_limits<unsigned __int64>::max() - inFlightTotalItems < total)
                {
                    inFlightTotalItems = std::numeric_limits<unsigned __int64>::max();
                }
                else
                {
                    inFlightTotalItems += total;
                }
            }

            const unsigned __int64 mappedCompletedBytes = _perItemCompletedBytes + inFlightCompletedBytes;
            _progressCompletedBytes                     = (std::max)(_progressCompletedBytes, mappedCompletedBytes);

            if (_operation == FILESYSTEM_DELETE)
            {
                const bool precalcTotalAvailable = _preCalcCompleted.load(std::memory_order_acquire) && _progressTotalItems > 0;
                if (! precalcTotalAvailable)
                {
                    const unsigned __int64 mappedTotalItems = _perItemTotalEntryCount + inFlightTotalItems;
                    if (mappedTotalItems > 0)
                    {
                        const unsigned __int64 clamped =
                            std::min<unsigned __int64>(mappedTotalItems, static_cast<unsigned __int64>(std::numeric_limits<unsigned long>::max()));
                        _progressTotalItems = (std::max)(_progressTotalItems, static_cast<unsigned long>(clamped));
                    }
                }

                const unsigned __int64 mappedCompletedItems = _perItemCompletedEntryCount + inFlightCompletedItems;
                const unsigned __int64 clamped =
                    std::min<unsigned __int64>(mappedCompletedItems, static_cast<unsigned __int64>(std::numeric_limits<unsigned long>::max()));
                _progressCompletedItems = (std::max)(_progressCompletedItems, static_cast<unsigned long>(clamped));
            }
            else
            {
                _progressCompletedItems = (std::max)(_progressCompletedItems, _perItemCompletedItems);
            }
        }
        else
        {
            if (totalItems > 0)
            {
                _progressTotalItems = (std::max)(_progressTotalItems, totalItems);
            }
            _progressCompletedItems = (std::max)(_progressCompletedItems, completedItems);
            if (totalBytes > 0)
            {
                _progressTotalBytes = (std::max)(_progressTotalBytes, totalBytes);
            }
            _progressCompletedBytes = (std::max)(_progressCompletedBytes, completedBytes);
        }

        if (_operation != FILESYSTEM_DELETE)
        {
            const unsigned long plannedTopLevelItems = (_executionMode == ExecutionMode::PerItem) ? _perItemTotalItems : GetPlannedItemCount();
            const bool havePreCalcTotals =
                _preCalcCompleted.load(std::memory_order_acquire) && _progressTotalItems > 0 && _progressTotalBytes > 0 && plannedTopLevelItems > 0;

            const bool pluginLikelyReportsTopLevelItems = (_executionMode == ExecutionMode::PerItem) ? true : (totalItems == 0 || totalItems <= plannedTopLevelItems);

            if (havePreCalcTotals && pluginLikelyReportsTopLevelItems && _progressTotalItems > plannedTopLevelItems)
            {
                const unsigned __int64 clampedBytes = (std::min)(_progressCompletedBytes, _progressTotalBytes);
                const long double ratio             = static_cast<long double>(clampedBytes) / static_cast<long double>(_progressTotalBytes);
                const long double estimate          = ratio * static_cast<long double>(_progressTotalItems);
                const long double clampedEstimate =
                    std::clamp<long double>(estimate, 0.0L, static_cast<long double>(_progressTotalItems));
                const unsigned long estimatedCompletedItems = static_cast<unsigned long>(clampedEstimate);
                _progressCompletedItems                     = (std::max)(_progressCompletedItems, estimatedCompletedItems);
            }
        }

        _progressItemTotalBytes     = currentItemTotalBytes;
        _progressItemCompletedBytes = currentItemCompletedBytes;

        _progressSourcePath                  = currentSourcePath ? currentSourcePath : L"";
        _progressDestinationPath             = currentDestinationPath ? currentDestinationPath : L"";
        _lastProgressCallbackSourcePath      = _progressSourcePath;
        _lastProgressCallbackDestinationPath = _progressDestinationPath;

        if (_executionMode == ExecutionMode::PerItem && cookie != nullptr)
        {
            auto* perItemCookie = static_cast<PerItemCallbackCookie*>(cookie);
            if (currentSourcePath && currentSourcePath[0] != L'\0')
            {
                perItemCookie->lastProgressSourcePath.assign(currentSourcePath);
            }
            if (currentDestinationPath && currentDestinationPath[0] != L'\0')
            {
                perItemCookie->lastProgressDestinationPath.assign(currentDestinationPath);
            }
        }

        if (options && (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE))
        {
            const unsigned __int64 pluginEffective = options->bandwidthLimitBytesPerSecond;
            const unsigned __int64 desiredTotal    = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);

            if (_executionMode == ExecutionMode::PerItem && _perItemMaxConcurrency > 1u)
            {
                unsigned __int64 desiredPerCall = desiredTotal;
                if (desiredTotal > 0)
                {
                    const unsigned int activeCalls = std::max(1u, static_cast<unsigned int>(_perItemInFlightCallCount));
                    desiredPerCall                 = std::max<unsigned __int64>(1ull, desiredTotal / static_cast<unsigned __int64>(activeCalls));
                }

                // Keep the UI limit line in task units (total), while applying the per-call share to the plugin.
                _effectiveSpeedLimitBytesPerSecond.store(desiredTotal, std::memory_order_release);
                options->bandwidthLimitBytesPerSecond = desiredPerCall;
                _appliedSpeedLimitBytesPerSecond.store(desiredPerCall, std::memory_order_release);
            }
            else
            {
                const unsigned __int64 applied = _appliedSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                _effectiveSpeedLimitBytesPerSecond.store(pluginEffective, std::memory_order_release);
                if (desiredTotal != applied)
                {
                    options->bandwidthLimitBytesPerSecond = desiredTotal;
                    _appliedSpeedLimitBytesPerSecond.store(desiredTotal, std::memory_order_release);
                }
            }
        }

        if ((_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE) && currentSourcePath && currentSourcePath[0] != L'\0')
        {
            // Keep a small in-flight set of file progress entries so the popup can display multiple file lines when the plugin runs in parallel.
            // Progress is tracked per (cookie, progressStreamId) so each concurrent worker can "own" a stable line and advance to new items.
            // Entries expire only when a new progress update arrives (so paused/waiting tasks keep their last view).
            constexpr ULONGLONG kExpiryMsActive    = 10'000ull;
            constexpr ULONGLONG kExpiryMsCompleted = 300ull;

            // Purge expired entries.
            size_t write = 0;
            for (size_t read = 0; read < _inFlightFileCount; ++read)
            {
                const InFlightFileProgress& entry = _inFlightFiles[read];
                const bool completed = entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes;
                const ULONGLONG expiryMs = completed ? kExpiryMsCompleted : kExpiryMsActive;
                const bool expired =
                    entry.lastUpdateTick != 0 && nowTick >= entry.lastUpdateTick && (nowTick - entry.lastUpdateTick) > expiryMs;
                if (expired)
                {
                    continue;
                }

                if (write != read)
                {
                    _inFlightFiles[write] = std::move(_inFlightFiles[read]);
                }
                ++write;
            }
            _inFlightFileCount = write;

            const void* cookieKey            = cookie;
            const unsigned __int64 streamKey = progressStreamId;

            // Find existing entry by (cookie, streamId).
            size_t found = _inFlightFileCount;
            for (size_t i = 0; i < _inFlightFileCount; ++i)
            {
                if (_inFlightFiles[i].cookieKey == cookieKey && _inFlightFiles[i].progressStreamId == streamKey)
                {
                    found = i;
                    break;
                }
            }

            if (found < _inFlightFileCount)
            {
                if (_inFlightFiles[found].sourcePath != currentSourcePath)
                {
                    _inFlightFiles[found].sourcePath.assign(currentSourcePath);
                }
                _inFlightFiles[found].totalBytes     = currentItemTotalBytes;
                _inFlightFiles[found].completedBytes = currentItemCompletedBytes;
                _inFlightFiles[found].lastUpdateTick = nowTick;
            }
            else
            {
                InFlightFileProgress added{};
                added.cookieKey         = cookieKey;
                added.progressStreamId  = streamKey;
                added.sourcePath     = currentSourcePath;
                added.totalBytes     = currentItemTotalBytes;
                added.completedBytes = currentItemCompletedBytes;
                added.lastUpdateTick = nowTick;

                if (_inFlightFileCount < _inFlightFiles.size())
                {
                    _inFlightFiles[_inFlightFileCount] = std::move(added);
                    found                              = _inFlightFileCount;
                    ++_inFlightFileCount;
                }
                else if (! _inFlightFiles.empty())
                {
                    // Replace the oldest entry (least recent update tick).
                    size_t replaceIndex = 0;
                    ULONGLONG oldestTick = _inFlightFiles[0].lastUpdateTick;
                    for (size_t i = 1; i < _inFlightFileCount; ++i)
                    {
                        const ULONGLONG tick = _inFlightFiles[i].lastUpdateTick;
                        if (tick == 0 || (oldestTick != 0 && tick < oldestTick))
                        {
                            replaceIndex = i;
                            oldestTick   = tick;
                        }
                    }

                    _inFlightFiles[replaceIndex] = std::move(added);
                    found                        = replaceIndex;
                }
            }
        }
    }

    WaitWhilePaused();

    if (_cancelled.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemItemCompleted(FileSystemOperation operationType,
                                                                                          unsigned long itemIndex,
                                                                                          const wchar_t* sourcePath,
                                                                                          const wchar_t* destinationPath,
                                                                                          HRESULT status,
                                                                                          FileSystemOptions* options,
                                                                                          void* cookie) noexcept
{
    if (operationType != _operation)
    {
        return S_OK;
    }

    {
        std::scoped_lock lock(_progressMutex);
        ++_itemCompletedCallbackCount;
        if (_executionMode != ExecutionMode::PerItem)
        {
            const unsigned long completedItemsClamped =
                static_cast<unsigned long>(std::min(_itemCompletedCallbackCount, static_cast<unsigned __int64>(ULONG_MAX)));
            _progressCompletedItems = (std::max)(_progressCompletedItems, completedItemsClamped);
        }
        _lastItemIndex           = itemIndex;
        _lastItemHr              = status;
        _progressSourcePath      = sourcePath ? sourcePath : L"";
        _progressDestinationPath = destinationPath ? destinationPath : L"";

        if (_executionMode == ExecutionMode::PerItem && cookie != nullptr)
        {
            auto* perItemCookie = static_cast<PerItemCallbackCookie*>(cookie);
            if (perItemCookie->lastProgressSourcePath.empty() && sourcePath && sourcePath[0] != L'\0')
            {
                perItemCookie->lastProgressSourcePath.assign(sourcePath);
            }
            if (perItemCookie->lastProgressDestinationPath.empty() && destinationPath && destinationPath[0] != L'\0')
            {
                perItemCookie->lastProgressDestinationPath.assign(destinationPath);
            }
        }

        // Best-effort cleanup when a top-level file item completes.
        if (sourcePath && sourcePath[0] != L'\0')
        {
            for (size_t i = 0; i < _inFlightFileCount; ++i)
            {
                if (_inFlightFiles[i].sourcePath == sourcePath)
                {
                    for (size_t j = i + 1u; j < _inFlightFileCount; ++j)
                    {
                        _inFlightFiles[j - 1u] = std::move(_inFlightFiles[j]);
                    }
                    --_inFlightFileCount;
                    break;
                }
            }
        }

        if (options && (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE))
        {
            const unsigned __int64 pluginEffective = options->bandwidthLimitBytesPerSecond;
            const unsigned __int64 desiredTotal    = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);

            if (_executionMode == ExecutionMode::PerItem && _perItemMaxConcurrency > 1u)
            {
                unsigned __int64 desiredPerCall = desiredTotal;
                if (desiredTotal > 0)
                {
                    const unsigned int activeCalls = std::max(1u, static_cast<unsigned int>(_perItemInFlightCallCount));
                    desiredPerCall                 = std::max<unsigned __int64>(1ull, desiredTotal / static_cast<unsigned __int64>(activeCalls));
                }

                _effectiveSpeedLimitBytesPerSecond.store(desiredTotal, std::memory_order_release);
                options->bandwidthLimitBytesPerSecond = desiredPerCall;
                _appliedSpeedLimitBytesPerSecond.store(desiredPerCall, std::memory_order_release);
            }
            else
            {
                const unsigned __int64 applied = _appliedSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                _effectiveSpeedLimitBytesPerSecond.store(pluginEffective, std::memory_order_release);
                if (desiredTotal != applied)
                {
                    options->bandwidthLimitBytesPerSecond = desiredTotal;
                    _appliedSpeedLimitBytesPerSecond.store(desiredTotal, std::memory_order_release);
                }
            }
        }
    }

    if (_cancelled.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemShouldCancel(BOOL* pCancel, void* /*cookie*/) noexcept
{
    if (! pCancel)
    {
        return E_POINTER;
    }

    const bool cancel = _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested();
    *pCancel          = cancel ? TRUE : FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemIssue(FileSystemOperation operationType,
                                                                                  const wchar_t* sourcePath,
                                                                                  const wchar_t* destinationPath,
                                                                                  HRESULT status,
                                                                                  FileSystemIssueAction* action,
                                                                                  [[maybe_unused]] FileSystemOptions* options,
                                                                                  void* cookie) noexcept
{
    if (! action)
    {
        return E_POINTER;
    }

    *action = FileSystemIssueAction::Cancel;

    WaitWhilePaused();

    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    const auto clearConflictPrompt = [&]() noexcept
    {
        {
            std::scoped_lock lock(_conflictMutex);
            _conflictPrompt = {};
            _conflictDecisionAction.reset();
            _conflictDecisionApplyToAll = false;
        }

        if (_conflictDecisionEvent)
        {
            static_cast<void>(ResetEvent(_conflictDecisionEvent.get()));
        }

        _conflictCv.notify_all();
    };

    const auto getMostSpecificPathsForDiagnostics = [&](const PerItemCallbackCookie* perItemCookie,
                                                        std::wstring_view sourceFallback,
                                                        std::wstring_view destinationFallback) noexcept -> std::pair<std::wstring, std::wstring>
    {
        std::wstring source(sourceFallback);
        std::wstring destination(destinationFallback);

        if (perItemCookie != nullptr)
        {
            std::scoped_lock lock(_progressMutex);
            if (! perItemCookie->lastProgressSourcePath.empty() &&
                (sourceFallback.empty() || IsSameOrChildPath(sourceFallback, perItemCookie->lastProgressSourcePath)))
            {
                source = perItemCookie->lastProgressSourcePath;
            }
            else if (! _lastProgressCallbackSourcePath.empty() &&
                     (sourceFallback.empty() || IsSameOrChildPath(sourceFallback, _lastProgressCallbackSourcePath)))
            {
                source = _lastProgressCallbackSourcePath;
            }

            if (! perItemCookie->lastProgressDestinationPath.empty() &&
                (destinationFallback.empty() || IsSameOrChildPath(destinationFallback, perItemCookie->lastProgressDestinationPath)))
            {
                destination = perItemCookie->lastProgressDestinationPath;
            }
            else if (! _lastProgressCallbackDestinationPath.empty() &&
                     (destinationFallback.empty() || IsSameOrChildPath(destinationFallback, _lastProgressCallbackDestinationPath)))
            {
                destination = _lastProgressCallbackDestinationPath;
            }
        }
        else
        {
            std::scoped_lock lock(_progressMutex);
            if (! _lastProgressCallbackSourcePath.empty() && (sourceFallback.empty() || IsSameOrChildPath(sourceFallback, _lastProgressCallbackSourcePath)))
            {
                source = _lastProgressCallbackSourcePath;
            }
            if (! _lastProgressCallbackDestinationPath.empty() &&
                (destinationFallback.empty() || IsSameOrChildPath(destinationFallback, _lastProgressCallbackDestinationPath)))
            {
                destination = _lastProgressCallbackDestinationPath;
            }
        }

        return {std::move(source), std::move(destination)};
    };

    const auto setCachedDecision = [&](ConflictBucket bucket, ConflictAction decision) noexcept
    {
        std::scoped_lock lock(_conflictMutex);
        _conflictDecisionCache[static_cast<size_t>(bucket)] = decision;
    };

    const auto getCachedDecision = [&](ConflictBucket bucket) noexcept -> std::optional<ConflictAction>
    {
        std::scoped_lock lock(_conflictMutex);
        return _conflictDecisionCache[static_cast<size_t>(bucket)];
    };

    const auto setConflictPromptLocked = [&](const PerItemCallbackCookie* perItemCookie,
                                             ConflictBucket bucket,
                                             HRESULT promptStatus,
                                             std::wstring_view sourceFallback,
                                             std::wstring_view destinationFallback,
                                             bool allowRetry,
                                             bool retryFailed) noexcept
    {
        auto [promptSourcePath, promptDestinationPath] = getMostSpecificPathsForDiagnostics(perItemCookie, sourceFallback, destinationFallback);

        auto addAction = [&](ConflictAction conflictAction) noexcept
        {
            if (_conflictPrompt.actionCount < _conflictPrompt.actions.size())
            {
                _conflictPrompt.actions[_conflictPrompt.actionCount] = conflictAction;
                ++_conflictPrompt.actionCount;
            }
        };

        if (_conflictDecisionEvent)
        {
            static_cast<void>(ResetEvent(_conflictDecisionEvent.get()));
        }

        _conflictPrompt                   = {};
        _conflictPrompt.active            = true;
        _conflictPrompt.bucket            = bucket;
        _conflictPrompt.status            = promptStatus;
        _conflictPrompt.sourcePath        = std::move(promptSourcePath);
        _conflictPrompt.destinationPath   = std::move(promptDestinationPath);
        _conflictPrompt.applyToAllChecked = false;
        _conflictPrompt.retryFailed       = retryFailed;
        _conflictPrompt.actionCount       = 0;

        LogDiagnostic(DiagnosticSeverity::Warning,
                      promptStatus,
                      L"item.conflict.prompt",
                      retryFailed ? L"Conflict prompt shown after retry cap reached." : L"Conflict prompt shown for item.",
                      _conflictPrompt.sourcePath,
                      _conflictPrompt.destinationPath);

        switch (bucket)
        {
            case ConflictBucket::Exists: addAction(ConflictAction::Overwrite); break;
            case ConflictBucket::ReadOnly: addAction(ConflictAction::ReplaceReadOnly); break;
            case ConflictBucket::RecycleBinFailed: addAction(ConflictAction::PermanentDelete); break;
            case ConflictBucket::AccessDenied:
            case ConflictBucket::SharingViolation:
            case ConflictBucket::DiskFull:
            case ConflictBucket::PathTooLong:
            case ConflictBucket::NetworkOffline:
            case ConflictBucket::UnsupportedReparse:
            case ConflictBucket::Unknown:
            case ConflictBucket::Count:
            default: break;
        }

        if (allowRetry)
        {
            addAction(ConflictAction::Retry);
        }

        addAction(ConflictAction::Skip);
        addAction(ConflictAction::Cancel);

        _conflictDecisionAction.reset();
        _conflictDecisionApplyToAll = false;
    };

    const auto waitForConflictDecision = [&]() noexcept -> std::pair<ConflictAction, bool>
    {
        if (! _conflictDecisionEvent)
        {
            clearConflictPrompt();
            return {ConflictAction::Cancel, false};
        }

        for (;;)
        {
            if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
            {
                clearConflictPrompt();
                return {ConflictAction::Cancel, false};
            }

            const DWORD wait = WaitForSingleObject(_conflictDecisionEvent.get(), 50);
            if (wait == WAIT_OBJECT_0)
            {
                break;
            }
        }

        ConflictAction decision = ConflictAction::Cancel;
        bool applyToAll         = false;
        {
            std::scoped_lock lock(_conflictMutex);
            decision   = _conflictDecisionAction.value_or(ConflictAction::Cancel);
            applyToAll = _conflictDecisionApplyToAll;
        }

        clearConflictPrompt();
        return {decision, applyToAll};
    };

    const std::wstring_view sourceText      = sourcePath ? sourcePath : L"";
    const std::wstring_view destinationText = destinationPath ? destinationPath : L"";

    PerItemCallbackCookie* perItemCookie = nullptr;
    if (_executionMode == ExecutionMode::PerItem && cookie != nullptr)
    {
        perItemCookie = static_cast<PerItemCallbackCookie*>(cookie);
    }

    const ConflictBucket bucket = ClassifyConflictBucket(operationType, _flags, wil::com_ptr<IFileSystemIO>{}, status, sourceText, destinationText, false);
    if (bucket == ConflictBucket::RecycleBinFailed)
    {
        auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(perItemCookie, sourceText, destinationText);
        LogDiagnostic(
            DiagnosticSeverity::Error, status, L"delete.recycleBin.item", L"Recycle Bin delete failed for item.", diagnosticSource, diagnosticDestination);
    }

    const size_t bucketIndex = static_cast<size_t>(bucket);

    ConflictAction decision = getCachedDecision(bucket).value_or(ConflictAction::None);
    if (decision == ConflictAction::None)
    {
        const bool canRetryBucket = bucket != ConflictBucket::UnsupportedReparse;
        bool allowRetry           = canRetryBucket;
        bool retryFailed          = false;
        if (perItemCookie != nullptr && bucketIndex < perItemCookie->issueRetryCounts.size())
        {
            allowRetry  = canRetryBucket && perItemCookie->issueRetryCounts[bucketIndex] == 0u;
            retryFailed = canRetryBucket && perItemCookie->issueRetryCounts[bucketIndex] != 0u;
        }

        {
            std::unique_lock lock(_conflictMutex);
            setConflictPromptLocked(perItemCookie, bucket, status, sourceText, destinationText, allowRetry, retryFailed);
        }

        const auto result     = waitForConflictDecision();
        decision              = result.first;
        const bool applyToAll = result.second;

        if (applyToAll && decision != ConflictAction::Retry && decision != ConflictAction::Cancel && decision != ConflictAction::None)
        {
            setCachedDecision(bucket, decision);
        }
    }

    switch (decision)
    {
        case ConflictAction::Overwrite: *action = FileSystemIssueAction::Overwrite; return S_OK;
        case ConflictAction::ReplaceReadOnly: *action = FileSystemIssueAction::ReplaceReadOnly; return S_OK;
        case ConflictAction::PermanentDelete: *action = FileSystemIssueAction::PermanentDelete; return S_OK;
        case ConflictAction::Retry:
            if (perItemCookie != nullptr && bucketIndex < perItemCookie->issueRetryCounts.size())
            {
                perItemCookie->issueRetryCounts[bucketIndex] = 1u;
            }
            *action = FileSystemIssueAction::Retry;
            return S_OK;
        case ConflictAction::SkipAll:
        {
            auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(perItemCookie, sourceText, destinationText);
            LogDiagnostic(DiagnosticSeverity::Warning,
                          status,
                          L"item.conflict.skipAll",
                          L"Conflict action Skip all similar conflicts selected.",
                          diagnosticSource,
                          diagnosticDestination);
            setCachedDecision(bucket, ConflictAction::Skip);
            *action = FileSystemIssueAction::Skip;
            return S_OK;
        }
        case ConflictAction::Skip:
        {
            auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(perItemCookie, sourceText, destinationText);
            LogDiagnostic(
                DiagnosticSeverity::Warning, status, L"item.conflict.skip", L"Conflict action Skip item selected.", diagnosticSource, diagnosticDestination);
            *action = FileSystemIssueAction::Skip;
            return S_OK;
        }
        case ConflictAction::Cancel:
        case ConflictAction::None:
        default: *action = FileSystemIssueAction::Cancel; return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::DirectorySizeProgress(unsigned __int64 /*scannedEntries*/,
                                                                                        unsigned __int64 totalBytes,
                                                                                        unsigned __int64 fileCount,
                                                                                        unsigned __int64 directoryCount,
                                                                                        const wchar_t* currentPath,
                                                                                        void* cookie) noexcept
{
    WaitWhilePreCalcPaused();

    const bool shouldCancel = _cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire) || _stopToken.stop_requested();
    if (shouldCancel)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    auto* progressCookie = static_cast<PreCalcProgressCookie*>(cookie);
    if (progressCookie && progressCookie->totalsMutex && progressCookie->totalBytes && progressCookie->totalFiles && progressCookie->totalDirs)
    {
        if (progressCookie->acceptUpdates && ! progressCookie->acceptUpdates->load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        const unsigned __int64 bytesDelta = (totalBytes >= progressCookie->lastBytes) ? (totalBytes - progressCookie->lastBytes) : totalBytes;
        const unsigned __int64 filesDelta = (fileCount >= progressCookie->lastFiles) ? (fileCount - progressCookie->lastFiles) : fileCount;
        const unsigned __int64 dirsDelta  = (directoryCount >= progressCookie->lastDirs) ? (directoryCount - progressCookie->lastDirs) : directoryCount;
        progressCookie->lastBytes         = totalBytes;
        progressCookie->lastFiles         = fileCount;
        progressCookie->lastDirs          = directoryCount;

        if (bytesDelta > 0 || filesDelta > 0 || dirsDelta > 0)
        {
            unsigned __int64 snapshotBytes = 0;
            unsigned __int64 snapshotFiles = 0;
            unsigned __int64 snapshotDirs  = 0;
            {
                std::scoped_lock lock(*progressCookie->totalsMutex);
                if (progressCookie->acceptUpdates && ! progressCookie->acceptUpdates->load(std::memory_order_acquire))
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                if (std::numeric_limits<unsigned __int64>::max() - *progressCookie->totalBytes < bytesDelta)
                {
                    *progressCookie->totalBytes = std::numeric_limits<unsigned __int64>::max();
                }
                else
                {
                    *progressCookie->totalBytes += bytesDelta;
                }
                if (std::numeric_limits<unsigned __int64>::max() - *progressCookie->totalFiles < filesDelta)
                {
                    *progressCookie->totalFiles = std::numeric_limits<unsigned __int64>::max();
                }
                else
                {
                    *progressCookie->totalFiles += filesDelta;
                }
                if (std::numeric_limits<unsigned __int64>::max() - *progressCookie->totalDirs < dirsDelta)
                {
                    *progressCookie->totalDirs = std::numeric_limits<unsigned __int64>::max();
                }
                else
                {
                    *progressCookie->totalDirs += dirsDelta;
                }

                snapshotBytes = *progressCookie->totalBytes;
                snapshotFiles = *progressCookie->totalFiles;
                snapshotDirs  = *progressCookie->totalDirs;
            }
            UpdatePreCalcSnapshot(*this, snapshotBytes, snapshotFiles, snapshotDirs);
        }
    }
    else
    {
        UpdatePreCalcSnapshot(*this, totalBytes, fileCount, directoryCount);
    }

    if (currentPath && currentPath[0] != L'\0')
    {
        std::scoped_lock lock(_progressMutex);
        _progressSourcePath = currentPath;
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::DirectorySizeShouldCancel(BOOL* pCancel, void* /*cookie*/) noexcept
{
    if (! pCancel)
    {
        return E_POINTER;
    }

    const bool cancel = _cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire) || _stopToken.stop_requested();
    *pCancel          = cancel ? TRUE : FALSE;
    return S_OK;
}

void FolderWindow::FileOperationState::Task::SkipPreCalculation() noexcept
{
    _preCalcSkipped.store(true, std::memory_order_release);
    LogDiagnostic(DiagnosticSeverity::Info, S_FALSE, L"precalc.skip", L"User skipped pre-calculation.");
    _pauseCv.notify_all();
}

void FolderWindow::FileOperationState::Task::RunPreCalculation() noexcept
{
    if (! _enablePreCalc || (_operation != FILESYSTEM_COPY && _operation != FILESYSTEM_MOVE && _operation != FILESYSTEM_DELETE) ||
        _preCalcSkipped.load(std::memory_order_acquire))
    {
        return;
    }

    if (_sourcePaths.empty())
    {
        return;
    }

    // Query IFileSystemDirectoryOperations interface
    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    if (FAILED(_fileSystem->QueryInterface(IID_PPV_ARGS(&dirOps))) || ! dirOps)
    {
        return; // Interface not supported, proceed without totals
    }

    _preCalcInProgress.store(true, std::memory_order_release);
    _preCalcStartTick.store(GetTickCount64(), std::memory_order_release);
    _preCalcCompleted.store(false, std::memory_order_release);
    _preCalcTotalBytes.store(0, std::memory_order_release);
    _preCalcFileCount.store(0, std::memory_order_release);
    _preCalcDirectoryCount.store(0, std::memory_order_release);

    _preCalcSourceBytes.clear();
    _preCalcSourceBytes.resize(_sourcePaths.size(), 0);

    std::mutex totalsMutex;
    unsigned __int64 totalBytes = 0;
    unsigned __int64 totalFiles = 0;
    unsigned __int64 totalDirs  = 0;
    std::atomic<bool> acceptUpdates{true};
    std::atomic<bool> preCalcAborted{false};

    const FileSystemFlags sizeFlags = FILESYSTEM_FLAG_RECURSIVE;

    const auto processIndex = [&](size_t index) noexcept
    {
        const auto& path = _sourcePaths[index];
        if (_cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire))
        {
            acceptUpdates.store(false, std::memory_order_release);
            preCalcAborted.store(true, std::memory_order_release);
            return;
        }

        PreCalcProgressCookie progressCookie{};
        progressCookie.totalsMutex   = &totalsMutex;
        progressCookie.totalBytes    = &totalBytes;
        progressCookie.totalFiles    = &totalFiles;
        progressCookie.totalDirs     = &totalDirs;
        progressCookie.acceptUpdates = &acceptUpdates;

        FileSystemDirectorySizeResult result{};
        const HRESULT hr     = dirOps->GetDirectorySize(path.c_str(), sizeFlags, this, &progressCookie, &result);
        const HRESULT status = FAILED(hr) ? hr : result.status;

        if (SUCCEEDED(status))
        {
            if (! acceptUpdates.load(std::memory_order_acquire))
            {
                return;
            }

            _preCalcSourceBytes[index] = result.totalBytes;

            const unsigned __int64 missingBytes =
                (result.totalBytes >= progressCookie.lastBytes) ? (result.totalBytes - progressCookie.lastBytes) : result.totalBytes;
            const unsigned __int64 missingFiles =
                (result.fileCount >= progressCookie.lastFiles) ? (result.fileCount - progressCookie.lastFiles) : result.fileCount;
            const unsigned __int64 missingDirs =
                (result.directoryCount >= progressCookie.lastDirs) ? (result.directoryCount - progressCookie.lastDirs) : result.directoryCount;

            if (missingBytes > 0 || missingFiles > 0 || missingDirs > 0)
            {
                unsigned __int64 snapshotBytes = 0;
                unsigned __int64 snapshotFiles = 0;
                unsigned __int64 snapshotDirs  = 0;
                {
                    std::scoped_lock lock(totalsMutex);
                    if (! acceptUpdates.load(std::memory_order_acquire))
                    {
                        return;
                    }

                    if (std::numeric_limits<unsigned __int64>::max() - totalBytes < missingBytes)
                    {
                        totalBytes = std::numeric_limits<unsigned __int64>::max();
                    }
                    else
                    {
                        totalBytes += missingBytes;
                    }
                    if (std::numeric_limits<unsigned __int64>::max() - totalFiles < missingFiles)
                    {
                        totalFiles = std::numeric_limits<unsigned __int64>::max();
                    }
                    else
                    {
                        totalFiles += missingFiles;
                    }
                    if (std::numeric_limits<unsigned __int64>::max() - totalDirs < missingDirs)
                    {
                        totalDirs = std::numeric_limits<unsigned __int64>::max();
                    }
                    else
                    {
                        totalDirs += missingDirs;
                    }

                    snapshotBytes = totalBytes;
                    snapshotFiles = totalFiles;
                    snapshotDirs  = totalDirs;
                }

                UpdatePreCalcSnapshot(*this, snapshotBytes, snapshotFiles, snapshotDirs);
            }
        }
        else if (status == HRESULT_FROM_WIN32(ERROR_CANCELLED))
        {
            acceptUpdates.store(false, std::memory_order_release);
            preCalcAborted.store(true, std::memory_order_release);
        }
        else
        {
            // Pre-calculation is best-effort, but failures are worth recording for debugging.
            const std::wstring statusText = FormatDiagnosticStatusText(status);
            LogDiagnostic(
                DiagnosticSeverity::Warning,
                status,
                L"precalc.error",
                std::format(L"Pre-calculation failed for '{}' (hr=0x{:08X}, status='{}').", path.c_str(), static_cast<unsigned long>(status), statusText),
                path.c_str());
        }
    };

    const size_t sourceCount = _sourcePaths.size();
    const bool useParallel   = sourceCount >= 2u;
    if (useParallel)
    {
        constexpr unsigned int kMaxPreCalcWorkers = 4u;
        const unsigned int workerCount            = static_cast<unsigned int>(std::min(sourceCount, static_cast<size_t>(kMaxPreCalcWorkers)));

        std::atomic<size_t> nextIndex{0};
        std::vector<std::jthread> workers;
        workers.reserve(workerCount);
        for (unsigned int worker = 0; worker < workerCount; ++worker)
        {
            workers.emplace_back(
                [&]() noexcept
                {
                    for (;;)
                    {
                        if (_cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire))
                        {
                            acceptUpdates.store(false, std::memory_order_release);
                            preCalcAborted.store(true, std::memory_order_release);
                            return;
                        }

                        const size_t index = nextIndex.fetch_add(1, std::memory_order_acq_rel);
                        if (index >= sourceCount)
                        {
                            return;
                        }

                        processIndex(index);
                    }
                });
        }
    }
    else
    {
        for (size_t index = 0; index < sourceCount; ++index)
        {
            processIndex(index);
            if (_cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire))
            {
                acceptUpdates.store(false, std::memory_order_release);
                preCalcAborted.store(true, std::memory_order_release);
                break;
            }
        }
    }

    _preCalcInProgress.store(false, std::memory_order_release);

    unsigned __int64 finalTotalBytes = 0;
    unsigned __int64 finalTotalFiles = 0;
    unsigned __int64 finalTotalDirs  = 0;
    {
        std::scoped_lock lock(totalsMutex);
        finalTotalBytes = totalBytes;
        finalTotalFiles = totalFiles;
        finalTotalDirs  = totalDirs;
    }
    UpdatePreCalcSnapshot(*this, finalTotalBytes, finalTotalFiles, finalTotalDirs);

    if (! _preCalcSkipped.load(std::memory_order_acquire) && ! _cancelled.load(std::memory_order_acquire) && ! preCalcAborted.load(std::memory_order_acquire))
    {
        _preCalcCompleted.store(true, std::memory_order_release);

        // Update progress totals if we got valid data
        if (finalTotalBytes > 0 || finalTotalFiles > 0 || finalTotalDirs > 0)
        {
            std::scoped_lock lock(_progressMutex);
            _progressTotalBytes = finalTotalBytes;
            _progressTotalItems = static_cast<unsigned long>(std::min(finalTotalFiles + finalTotalDirs, static_cast<unsigned __int64>(ULONG_MAX)));
        }
    }
}

void FolderWindow::FileOperationState::Task::ThreadMain(std::stop_token stopToken) noexcept
{
    _stopToken                   = stopToken;
    [[maybe_unused]] auto coInit = wil::CoInitializeEx();
    [[maybe_unused]] const std::stop_callback stopWake(stopToken,
                                                       [this]() noexcept
                                                       {
                                                           _pauseCv.notify_all();
                                                           _conflictCv.notify_all();
                                                           if (_state)
                                                           {
                                                               _state->NotifyQueueChanged();
                                                           }
                                                       });

    if (! _state)
    {
        return;
    }

    LogDiagnostic(DiagnosticSeverity::Debug,
                  S_OK,
                  L"task.started",
                  std::format(L"Task started (op={}, mode={}, sources={}, flags=0x{:08X}, preCalc={}, waitForOthers={}).",
                              OperationToString(_operation),
                              _executionMode == ExecutionMode::PerItem ? L"perItem" : L"bulkItems",
                              _sourcePaths.size(),
                              static_cast<unsigned long>(static_cast<uint32_t>(_flags)),
                              _enablePreCalc ? L"on" : L"off",
                              _waitForOthers.load(std::memory_order_acquire) ? L"true" : L"false"));

    // Mark as waiting in queue before entering (visible to UI while blocked). Use the current
    // desired start-gating state to avoid briefly showing "Waiting" for tasks that will start immediately.
    _waitingInQueue.store(_waitForOthers.load(std::memory_order_acquire), std::memory_order_release);

    // Enter queue FIRST so both pre-calculation and operation respect Wait/Parallel mode
    const bool canStart = _state->EnterOperation(*this, stopToken);

    // No longer waiting in queue (either we got our turn or were cancelled)
    _waitingInQueue.store(false, std::memory_order_release);

    if (! canStart)
    {
        _resultHr.store(HRESULT_FROM_WIN32(ERROR_CANCELLED), std::memory_order_release);
        _state->PostCompleted(*this);
        return;
    }

    _enteredOperationTick.store(GetTickCount64(), std::memory_order_release);
    _enteredOperation.store(true, std::memory_order_release);

    // Run pre-calculation phase while holding queue slot
    RunPreCalculation();

    const ULONGLONG afterPreCalcTick = GetTickCount64();
    if (Debug::Perf::IsEnabled())
    {
        const ULONGLONG preStartTick = _preCalcStartTick.load(std::memory_order_acquire);
        if (preStartTick > 0)
        {
            const ULONGLONG elapsedMs    = (afterPreCalcTick >= preStartTick) ? (afterPreCalcTick - preStartTick) : 0;
            const uint64_t durationUs    = static_cast<uint64_t>(elapsedMs) * 1000ull;
            const HRESULT preCalcHr      = _cancelled.load(std::memory_order_acquire) ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                                                      : (_preCalcSkipped.load(std::memory_order_acquire) ? S_FALSE : S_OK);
            const unsigned __int64 bytes = _preCalcTotalBytes.load(std::memory_order_acquire);
            const unsigned __int64 items = static_cast<unsigned __int64>(_preCalcFileCount.load(std::memory_order_acquire)) +
                                           static_cast<unsigned __int64>(_preCalcDirectoryCount.load(std::memory_order_acquire));

            const size_t sourceCount  = _sourcePaths.size();
            const std::wstring detail = std::format(L"id={} op={} sources={}", _taskId, OperationToString(_operation), sourceCount);
            Debug::Perf::Emit(L"FileOps.PreCalc", detail, durationUs, bytes, items, preCalcHr);
        }
    }

    {
        const ULONGLONG preStartTick = _preCalcStartTick.load(std::memory_order_acquire);
        if (preStartTick > 0)
        {
            const ULONGLONG elapsedMs    = (afterPreCalcTick >= preStartTick) ? (afterPreCalcTick - preStartTick) : 0;
            const HRESULT preCalcHr      = _cancelled.load(std::memory_order_acquire) ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                                                      : (_preCalcSkipped.load(std::memory_order_acquire) ? S_FALSE : S_OK);
            const unsigned __int64 bytes = _preCalcTotalBytes.load(std::memory_order_acquire);
            const unsigned long files    = _preCalcFileCount.load(std::memory_order_acquire);
            const unsigned long dirs     = _preCalcDirectoryCount.load(std::memory_order_acquire);
            const bool skipped           = _preCalcSkipped.load(std::memory_order_acquire);
            LogDiagnostic(DiagnosticSeverity::Debug,
                          preCalcHr,
                          L"precalc.result",
                          std::format(L"Pre-calculation finished (hr=0x{:08X}, elapsedMs={}, bytes={:L}, files={:L}, dirs={:L}, skipped={}).",
                                      static_cast<unsigned long>(preCalcHr),
                                      elapsedMs,
                                      bytes,
                                      files,
                                      dirs,
                                      skipped ? L"true" : L"false"));
        }
    }

    // Check if cancelled during pre-calc
    if (_cancelled.load(std::memory_order_acquire))
    {
        _enteredOperation.store(false, std::memory_order_release);
        _enteredOperationTick.store(0, std::memory_order_release);
        _state->LeaveOperation();
        _resultHr.store(HRESULT_FROM_WIN32(ERROR_CANCELLED), std::memory_order_release);
        _state->PostCompleted(*this);
        return;
    }

    const HRESULT hr = ExecuteOperation();
    _resultHr.store(hr, std::memory_order_release);

    if (FAILED(hr))
    {
        unsigned long totalItems        = 0;
        unsigned long completedItems    = 0;
        unsigned __int64 totalBytes     = 0;
        unsigned __int64 completedBytes = 0;
        std::wstring sourcePath;
        std::wstring destinationPath;
        {
            std::scoped_lock lock(_progressMutex);
            totalItems      = _progressTotalItems;
            completedItems  = _progressCompletedItems;
            totalBytes      = _progressTotalBytes;
            completedBytes  = _progressCompletedBytes;
            sourcePath      = _progressSourcePath;
            destinationPath = _progressDestinationPath;
        }

        const HRESULT partialCopyHr       = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        const DiagnosticSeverity severity = (hr == partialCopyHr)      ? DiagnosticSeverity::Warning
                                            : IsCancellationStatus(hr) ? DiagnosticSeverity::Info
                                                                       : DiagnosticSeverity::Error;
        std::wstring message;
        if (hr == partialCopyHr)
        {
            message = std::format(L"Task completed with skipped or partial items (op={}, items={:L}/{:L}, bytes={:L}/{:L}).",
                                  OperationToString(_operation),
                                  completedItems,
                                  totalItems,
                                  completedBytes,
                                  totalBytes);
        }
        else if (IsCancellationStatus(hr))
        {
            message = std::format(L"Task was canceled (op={}, items={:L}/{:L}, bytes={:L}/{:L}).",
                                  OperationToString(_operation),
                                  completedItems,
                                  totalItems,
                                  completedBytes,
                                  totalBytes);
        }
        else
        {
            const std::wstring statusText = FormatDiagnosticStatusText(hr);
            message                       = std::format(L"Task failed (op={}, hr=0x{:08X}, status='{}', items={:L}/{:L}, bytes={:L}/{:L}).",
                                  OperationToString(_operation),
                                  static_cast<unsigned long>(hr),
                                  statusText,
                                  completedItems,
                                  totalItems,
                                  completedBytes,
                                  totalBytes);
        }
        LogDiagnostic(severity, hr, L"task.result", message, sourcePath, destinationPath);
    }

    {
        const ULONGLONG opStartTick = _operationStartTick.load(std::memory_order_acquire);
        const ULONGLONG endTick     = GetTickCount64();
        const ULONGLONG elapsedMs   = (opStartTick > 0 && endTick >= opStartTick) ? (endTick - opStartTick) : 0;

        unsigned long totalItems        = 0;
        unsigned long completedItems    = 0;
        unsigned __int64 totalBytes     = 0;
        unsigned __int64 completedBytes = 0;
        unsigned __int64 progressCalls  = 0;
        unsigned __int64 itemCalls      = 0;
        std::wstring sourcePath;
        std::wstring destinationPath;
        {
            std::scoped_lock lock(_progressMutex);
            totalItems      = _progressTotalItems;
            completedItems  = _progressCompletedItems;
            totalBytes      = _progressTotalBytes;
            completedBytes  = _progressCompletedBytes;
            progressCalls   = _progressCallbackCount;
            itemCalls       = _itemCompletedCallbackCount;
            sourcePath      = _progressSourcePath;
            destinationPath = _progressDestinationPath;
        }

        LogDiagnostic(DiagnosticSeverity::Debug,
                      hr,
                      L"task.operation.result",
                      std::format(L"Operation finished (hr=0x{:08X}, elapsedMs={}, items={:L}/{:L}, bytes={:L}/{:L}, progressCalls={:L}, itemCalls={:L}).",
                                  static_cast<unsigned long>(hr),
                                  elapsedMs,
                                  completedItems,
                                  totalItems,
                                  completedBytes,
                                  totalBytes,
                                  progressCalls,
                                  itemCalls),
                      sourcePath,
                      destinationPath);
    }

    if (Debug::Perf::IsEnabled())
    {
        const ULONGLONG opStartTick = _operationStartTick.load(std::memory_order_acquire);
        const ULONGLONG endTick     = GetTickCount64();
        const ULONGLONG elapsedMs   = (opStartTick > 0 && endTick >= opStartTick) ? (endTick - opStartTick) : 0;
        const uint64_t durationUs   = static_cast<uint64_t>(elapsedMs) * 1000ull;

        unsigned __int64 completedBytes = 0;
        unsigned long completedItems    = 0;
        unsigned __int64 progressCalls  = 0;
        unsigned __int64 itemCalls      = 0;
        {
            std::scoped_lock lock(_progressMutex);
            completedBytes = _progressCompletedBytes;
            completedItems = _progressCompletedItems;
            progressCalls  = _progressCallbackCount;
            itemCalls      = _itemCompletedCallbackCount;
        }

        const unsigned __int64 desired   = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        const unsigned __int64 effective = _effectiveSpeedLimitBytesPerSecond.load(std::memory_order_acquire);

        const size_t sourceCount  = _sourcePaths.size();
        const std::wstring detail = std::format(L"id={} op={} desired={} effective={} sources={} items={}",
                                                _taskId,
                                                OperationToString(_operation),
                                                desired,
                                                effective,
                                                sourceCount,
                                                completedItems);
        Debug::Perf::Emit(L"FileOps.Operation", detail, durationUs, completedBytes, progressCalls, hr);

        const ULONGLONG cancelTick = _cancelRequestedTick.load(std::memory_order_acquire);
        if (cancelTick > 0)
        {
            const ULONGLONG cancelMs = (endTick >= cancelTick) ? (endTick - cancelTick) : 0;
            const uint64_t cancelUs  = static_cast<uint64_t>(cancelMs) * 1000ull;
            Debug::Perf::Emit(L"FileOps.CancelLatency", detail, cancelUs, completedBytes, itemCalls, hr);
        }
    }

    _enteredOperation.store(false, std::memory_order_release);
    _enteredOperationTick.store(0, std::memory_order_release);
    _state->LeaveOperation();
    _state->PostCompleted(*this);
}

void FolderWindow::FileOperationState::Task::RequestCancel() noexcept
{
    {
        ULONGLONG expected = 0;
        _cancelRequestedTick.compare_exchange_strong(expected, GetTickCount64(), std::memory_order_release);
    }
    _cancelled.store(true, std::memory_order_release);
    {
        std::scoped_lock lock(_pauseMutex);
        _paused.store(false, std::memory_order_release);
    }
    _pauseCv.notify_all();

    if (_conflictDecisionEvent)
    {
        static_cast<void>(SetEvent(_conflictDecisionEvent.get()));
    }

    _conflictCv.notify_all();

    if (_state)
    {
        _state->NotifyQueueChanged();
    }
}

void FolderWindow::FileOperationState::Task::TogglePause() noexcept
{
    const bool nowPaused = ! _paused.load(std::memory_order_acquire);
    _paused.store(nowPaused, std::memory_order_release);
    if (! nowPaused)
    {
        _pauseCv.notify_all();
    }
}

void FolderWindow::FileOperationState::Task::SetDesiredSpeedLimit(unsigned __int64 bytesPerSecond) noexcept
{
    _desiredSpeedLimitBytesPerSecond.store(bytesPerSecond, std::memory_order_release);
}

void FolderWindow::FileOperationState::Task::SetWaitForOthers(bool wait) noexcept
{
    if (_started.load(std::memory_order_acquire))
    {
        return;
    }

    _waitForOthers.store(wait, std::memory_order_release);
    if (_state)
    {
        _state->NotifyQueueChanged();
    }
}

void FolderWindow::FileOperationState::Task::SetQueuePaused(bool paused) noexcept
{
    const bool wasPaused = _queuePaused.load(std::memory_order_acquire);
    if (wasPaused == paused)
    {
        return;
    }

    _queuePaused.store(paused, std::memory_order_release);
    if (! paused)
    {
        _pauseCv.notify_all();
    }
}

void FolderWindow::FileOperationState::Task::ToggleConflictApplyToAllChecked() noexcept
{
    std::scoped_lock lock(_conflictMutex);
    if (! _conflictPrompt.active)
    {
        return;
    }

    _conflictPrompt.applyToAllChecked = ! _conflictPrompt.applyToAllChecked;
}

void FolderWindow::FileOperationState::Task::SubmitConflictDecision(ConflictAction action, bool applyToAllChecked) noexcept
{
    {
        std::scoped_lock lock(_conflictMutex);
        if (! _conflictPrompt.active)
        {
            return;
        }

        _conflictDecisionAction     = action;
        _conflictDecisionApplyToAll = (action == ConflictAction::Retry) ? false : (applyToAllChecked || action == ConflictAction::SkipAll);
    }

    if (_conflictDecisionEvent)
    {
        static_cast<void>(SetEvent(_conflictDecisionEvent.get()));
    }
}

bool FolderWindow::FileOperationState::Task::HasStarted() const noexcept
{
    return _started.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::Task::HasEnteredOperation() const noexcept
{
    return _enteredOperation.load(std::memory_order_acquire);
}

ULONGLONG FolderWindow::FileOperationState::Task::GetEnteredOperationTick() const noexcept
{
    return _enteredOperationTick.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::Task::IsPaused() const noexcept
{
    return _paused.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::Task::IsWaitingForOthers() const noexcept
{
    return _waitForOthers.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::Task::IsWaitingInQueue() const noexcept
{
    return _waitingInQueue.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::Task::IsQueuePaused() const noexcept
{
    return _queuePaused.load(std::memory_order_acquire);
}

void FolderWindow::FileOperationState::Task::SetDestinationFolder(const std::filesystem::path& folder)
{
    if (_started.load(std::memory_order_acquire))
    {
        return;
    }

    std::scoped_lock lock(_operationMutex);
    _destinationFolder = folder;
}

std::filesystem::path FolderWindow::FileOperationState::Task::GetDestinationFolder() const
{
    std::scoped_lock lock(_operationMutex);
    return _destinationFolder;
}

unsigned long FolderWindow::FileOperationState::Task::GetPlannedItemCount() const noexcept
{
    const unsigned __int64 count64 = static_cast<unsigned __int64>(_sourcePaths.size());
    if (count64 > std::numeric_limits<unsigned long>::max())
    {
        return std::numeric_limits<unsigned long>::max();
    }
    return static_cast<unsigned long>(count64);
}

uint64_t FolderWindow::FileOperationState::Task::GetId() const noexcept
{
    return _taskId;
}

HRESULT FolderWindow::FileOperationState::Task::GetResult() const noexcept
{
    return _resultHr.load(std::memory_order_acquire);
}

FileSystemOperation FolderWindow::FileOperationState::Task::GetOperation() const noexcept
{
    return _operation;
}

FolderWindow::Pane FolderWindow::FileOperationState::Task::GetSourcePane() const noexcept
{
    return _sourcePane;
}

std::optional<FolderWindow::Pane> FolderWindow::FileOperationState::Task::GetDestinationPane() const noexcept
{
    return _destinationPane;
}

void FolderWindow::FileOperationState::Task::WaitWhilePaused() noexcept
{
    const bool shouldPause = _paused.load(std::memory_order_acquire) || _queuePaused.load(std::memory_order_acquire);
    if (! shouldPause)
    {
        return;
    }

    std::unique_lock lock(_pauseMutex);
    _pauseCv.wait(lock,
                  [&]
                  {
                      const bool stillPaused = _paused.load(std::memory_order_acquire) || _queuePaused.load(std::memory_order_acquire);
                      return ! stillPaused || _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested();
                  });
}

void FolderWindow::FileOperationState::Task::WaitWhilePreCalcPaused() noexcept
{
    const bool shouldPause = _paused.load(std::memory_order_acquire) || _queuePaused.load(std::memory_order_acquire);
    if (! shouldPause)
    {
        return;
    }

    std::unique_lock lock(_pauseMutex);
    _pauseCv.wait(lock,
                  [&]
                  {
                      const bool stillPaused = _paused.load(std::memory_order_acquire) || _queuePaused.load(std::memory_order_acquire);
                      return ! stillPaused || _cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire) ||
                             _stopToken.stop_requested();
                  });
}

HRESULT FolderWindow::FileOperationState::Task::ExecuteOperation() noexcept
{
    if (! _fileSystem)
    {
        return E_POINTER;
    }

    if (_sourcePaths.empty())
    {
        return S_FALSE;
    }

    WaitWhilePaused();
    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    _started.store(true, std::memory_order_release);
    _operationStartTick.store(GetTickCount64(), std::memory_order_release);

    std::filesystem::path destinationFolder;
    {
        std::scoped_lock lock(_operationMutex);
        destinationFolder = _destinationFolder;
    }

    const bool continueOnError = (_flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;

    if (_executionMode == ExecutionMode::PerItem)
    {
        wil::com_ptr<IFileSystemIO> fileSystemIo;
        static_cast<void>(_fileSystem->QueryInterface(IID_PPV_ARGS(fileSystemIo.addressof())));

        const bool useCrossFileSystemBridge = (_destinationFileSystem != nullptr) && (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE);

        wil::com_ptr<IFileSystemIO> destinationFileSystemIo;
        wil::com_ptr<IFileSystemDirectoryOperations> destinationDirOps;
        if (useCrossFileSystemBridge)
        {
            static_cast<void>(_destinationFileSystem->QueryInterface(IID_PPV_ARGS(destinationFileSystemIo.addressof())));
            static_cast<void>(_destinationFileSystem->QueryInterface(IID_PPV_ARGS(destinationDirOps.addressof())));

            if (! fileSystemIo || ! destinationFileSystemIo)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }
        }

        ReparsePointPolicy reparsePointPolicy = ReparsePointPolicy::CopyReparse;
        if (_folderWindow && _folderWindow->_settings)
        {
            const std::wstring& sourcePluginId =
                _sourcePane == FolderWindow::Pane::Left ? _folderWindow->_leftPane.pluginId : _folderWindow->_rightPane.pluginId;
            if (! sourcePluginId.empty())
            {
                reparsePointPolicy = GetReparsePointPolicyFromSettings(*_folderWindow->_settings, sourcePluginId);
            }
        }

        const unsigned __int64 count64 = static_cast<unsigned __int64>(_sourcePaths.size());
        if (count64 > std::numeric_limits<unsigned long>::max())
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        _perItemTotalItems     = static_cast<unsigned long>(count64);
        _perItemMaxConcurrency = DeterminePerItemMaxConcurrency(_fileSystem, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles));
        _perItemMaxConcurrency = std::max(1u, _perItemMaxConcurrency);
        _perItemMaxConcurrency = std::min<unsigned int>(_perItemMaxConcurrency, static_cast<unsigned int>(_perItemTotalItems));
        if (useCrossFileSystemBridge)
        {
            const unsigned int destinationMaxConcurrency =
                DeterminePerItemMaxConcurrency(_destinationFileSystem, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles));
            _perItemMaxConcurrency = std::min(_perItemMaxConcurrency, destinationMaxConcurrency);
            _perItemMaxConcurrency = std::max(1u, _perItemMaxConcurrency);
        }
        _perItemCompletedItems      = 0;
        _perItemCompletedEntryCount = 0;
        _perItemTotalEntryCount     = 0;
        _perItemCompletedBytes      = 0;
        _perItemInFlightCallCount   = 0;

        {
            std::scoped_lock lock(_progressMutex);
            if (_operation != FILESYSTEM_DELETE)
            {
                _progressTotalItems = _perItemTotalItems;
            }
            _progressCompletedItems = 0;
            _progressCompletedBytes = 0;
        }

        const bool canUsePreCalcBytes = _preCalcCompleted.load(std::memory_order_acquire) && _preCalcSourceBytes.size() == _sourcePaths.size();

        bool hadSkippedItems = false;

        if ((_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE) && destinationFolder.empty())
        {
            return E_INVALIDARG;
        }

        const std::wstring destinationFolderText = destinationFolder.native();

        const auto clearConflictPrompt = [&]() noexcept
        {
            {
                std::scoped_lock lock(_conflictMutex);
                _conflictPrompt = {};
                _conflictDecisionAction.reset();
                _conflictDecisionApplyToAll = false;
            }

            if (_conflictDecisionEvent)
            {
                static_cast<void>(ResetEvent(_conflictDecisionEvent.get()));
            }

            _conflictCv.notify_all();
        };

        const auto getMostSpecificPathsForDiagnostics = [&](const PerItemCallbackCookie* perItemCookie,
                                                            std::wstring_view sourceFallback,
                                                            std::wstring_view destinationFallback) noexcept -> std::pair<std::wstring, std::wstring>
        {
            std::wstring source(sourceFallback);
            std::wstring destination(destinationFallback);

            if (perItemCookie != nullptr)
            {
                if (! perItemCookie->lastProgressSourcePath.empty() &&
                    (sourceFallback.empty() || IsSameOrChildPath(sourceFallback, perItemCookie->lastProgressSourcePath)))
                {
                    source = perItemCookie->lastProgressSourcePath;
                }
                if (! perItemCookie->lastProgressDestinationPath.empty() &&
                    (destinationFallback.empty() || IsSameOrChildPath(destinationFallback, perItemCookie->lastProgressDestinationPath)))
                {
                    destination = perItemCookie->lastProgressDestinationPath;
                }

                return {std::move(source), std::move(destination)};
            }

            {
                std::scoped_lock lock(_progressMutex);
                if (! _lastProgressCallbackSourcePath.empty() && (sourceFallback.empty() || IsSameOrChildPath(sourceFallback, _lastProgressCallbackSourcePath)))
                {
                    source = _lastProgressCallbackSourcePath;
                }
                else if (! _progressSourcePath.empty() && (sourceFallback.empty() || IsSameOrChildPath(sourceFallback, _progressSourcePath)))
                {
                    source = _progressSourcePath;
                }
                if (! _lastProgressCallbackDestinationPath.empty() &&
                    (destinationFallback.empty() || IsSameOrChildPath(destinationFallback, _lastProgressCallbackDestinationPath)))
                {
                    destination = _lastProgressCallbackDestinationPath;
                }
                else if (! _progressDestinationPath.empty() &&
                         (destinationFallback.empty() || IsSameOrChildPath(destinationFallback, _progressDestinationPath)))
                {
                    destination = _progressDestinationPath;
                }
            }
            return {std::move(source), std::move(destination)};
        };

        const auto setConflictPromptLocked = [&](const PerItemCallbackCookie* perItemCookie,
                                                 ConflictBucket bucket,
                                                 HRESULT status,
                                                 std::wstring_view sourcePath,
                                                 std::wstring_view destinationPath,
                                                 bool allowRetry,
                                                 bool retryFailed) noexcept
        {
            auto [promptSourcePath, promptDestinationPath] = getMostSpecificPathsForDiagnostics(perItemCookie, sourcePath, destinationPath);

            auto addAction = [&](ConflictAction action) noexcept
            {
                if (_conflictPrompt.actionCount < _conflictPrompt.actions.size())
                {
                    _conflictPrompt.actions[_conflictPrompt.actionCount] = action;
                    ++_conflictPrompt.actionCount;
                }
            };

            if (_conflictDecisionEvent)
            {
                static_cast<void>(ResetEvent(_conflictDecisionEvent.get()));
            }
            _conflictPrompt                   = {};
            _conflictPrompt.active            = true;
            _conflictPrompt.bucket            = bucket;
            _conflictPrompt.status            = status;
            _conflictPrompt.sourcePath        = std::move(promptSourcePath);
            _conflictPrompt.destinationPath   = std::move(promptDestinationPath);
            _conflictPrompt.applyToAllChecked = false;
            _conflictPrompt.retryFailed       = retryFailed;

            _conflictPrompt.actionCount = 0;

            LogDiagnostic(DiagnosticSeverity::Warning,
                          status,
                          L"item.conflict.prompt",
                          retryFailed ? L"Conflict prompt shown after retry cap reached." : L"Conflict prompt shown for item.",
                          _conflictPrompt.sourcePath,
                          _conflictPrompt.destinationPath);

            switch (bucket)
            {
                case ConflictBucket::Exists: addAction(ConflictAction::Overwrite); break;
                case ConflictBucket::ReadOnly: addAction(ConflictAction::ReplaceReadOnly); break;
                case ConflictBucket::RecycleBinFailed: addAction(ConflictAction::PermanentDelete); break;
                case ConflictBucket::AccessDenied:
                case ConflictBucket::SharingViolation:
                case ConflictBucket::DiskFull:
                case ConflictBucket::PathTooLong:
                case ConflictBucket::NetworkOffline:
                case ConflictBucket::UnsupportedReparse:
                case ConflictBucket::Unknown:
                case ConflictBucket::Count:
                default: break;
            }

            if (allowRetry)
            {
                addAction(ConflictAction::Retry);
            }
            addAction(ConflictAction::Skip);
            addAction(ConflictAction::Cancel);

            _conflictDecisionAction.reset();
            _conflictDecisionApplyToAll = false;
        };

        const auto waitForConflictDecision = [&]() noexcept -> std::pair<ConflictAction, bool>
        {
            if (! _conflictDecisionEvent)
            {
                clearConflictPrompt();
                return {ConflictAction::Cancel, false};
            }

            for (;;)
            {
                if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
                {
                    clearConflictPrompt();
                    return {ConflictAction::Cancel, false};
                }

                const DWORD wait = WaitForSingleObject(_conflictDecisionEvent.get(), 50);
                if (wait == WAIT_OBJECT_0)
                {
                    break;
                }
            }

            ConflictAction action = ConflictAction::Cancel;
            bool applyToAll       = false;
            {
                std::scoped_lock lock(_conflictMutex);
                action     = _conflictDecisionAction.value_or(ConflictAction::Cancel);
                applyToAll = _conflictDecisionApplyToAll;
            }

            clearConflictPrompt();
            return {action, applyToAll};
        };

        const auto getCachedDecision = [&](ConflictBucket bucket) noexcept -> std::optional<ConflictAction>
        {
            std::scoped_lock lock(_conflictMutex);
            return _conflictDecisionCache[static_cast<size_t>(bucket)];
        };

        const auto setCachedDecision = [&](ConflictBucket bucket, ConflictAction action) noexcept
        {
            if (action == ConflictAction::Retry || action == ConflictAction::Cancel || action == ConflictAction::None)
            {
                return;
            }

            if (action == ConflictAction::SkipAll)
            {
                action = ConflictAction::Skip;
            }

            std::scoped_lock lock(_conflictMutex);
            _conflictDecisionCache[static_cast<size_t>(bucket)] = action;
        };

        const auto clearCachedDecision = [&](ConflictBucket bucket) noexcept
        {
            std::scoped_lock lock(_conflictMutex);
            _conflictDecisionCache[static_cast<size_t>(bucket)].reset();
        };

        const auto isModifierConflictAction = [](ConflictAction action) noexcept
        {
            switch (action)
            {
                case ConflictAction::Overwrite:
                case ConflictAction::ReplaceReadOnly:
                case ConflictAction::PermanentDelete: return true;
                case ConflictAction::None:
                case ConflictAction::Retry:
                case ConflictAction::Skip:
                case ConflictAction::SkipAll:
                case ConflictAction::Cancel:
                default: return false;
            }
        };

        constexpr unsigned int kMaxCachedModifierAttemptsPerBucket = 1u;

        const auto computeInFlightCompletedBytesLocked = [&]() noexcept -> unsigned __int64
        {
            unsigned __int64 sum = 0;
            for (size_t i = 0; i < _perItemInFlightCallCount; ++i)
            {
                const unsigned __int64 v = _perItemInFlightCalls[i].completedBytes;
                if (std::numeric_limits<unsigned __int64>::max() - sum < v)
                {
                    return std::numeric_limits<unsigned __int64>::max();
                }
                sum += v;
            }
            return sum;
        };

        const auto computeInFlightCompletedItemsLocked = [&]() noexcept -> unsigned __int64
        {
            unsigned __int64 sum = 0;
            for (size_t i = 0; i < _perItemInFlightCallCount; ++i)
            {
                const unsigned __int64 v = static_cast<unsigned __int64>(_perItemInFlightCalls[i].completedItems);
                if (std::numeric_limits<unsigned __int64>::max() - sum < v)
                {
                    return std::numeric_limits<unsigned __int64>::max();
                }
                sum += v;
            }
            return sum;
        };

        const auto computeInFlightTotalItemsLocked = [&]() noexcept -> unsigned __int64
        {
            unsigned __int64 sum = 0;
            for (size_t i = 0; i < _perItemInFlightCallCount; ++i)
            {
                const unsigned __int64 v = static_cast<unsigned __int64>(_perItemInFlightCalls[i].totalItems);
                if (std::numeric_limits<unsigned __int64>::max() - sum < v)
                {
                    return std::numeric_limits<unsigned __int64>::max();
                }
                sum += v;
            }
            return sum;
        };

        struct BridgeCallback final : IFileSystemCallback
        {
            Task& task;

            explicit BridgeCallback(Task& owner) noexcept : task(owner)
            {
            }

            BridgeCallback(const BridgeCallback&)            = delete;
            BridgeCallback(BridgeCallback&&)                 = delete;
            BridgeCallback& operator=(const BridgeCallback&) = delete;
            BridgeCallback& operator=(BridgeCallback&&)      = delete;

             HRESULT STDMETHODCALLTYPE FileSystemProgress(FileSystemOperation /*operationType*/,
                                                          unsigned long /*totalItems*/,
                                                          unsigned long /*completedItems*/,
                                                          unsigned __int64 /*totalBytes*/,
                                                          unsigned __int64 /*completedBytes*/,
                                                          const wchar_t* /*currentSourcePath*/,
                                                          const wchar_t* /*currentDestinationPath*/,
                                                          unsigned __int64 /*currentItemTotalBytes*/,
                                                          unsigned __int64 /*currentItemCompletedBytes*/,
                                                          FileSystemOptions* /*options*/,
                                                          unsigned __int64 /*progressStreamId*/,
                                                          void* /*cookie*/) noexcept override
             {
                 task.WaitWhilePaused();
                 if (task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested())
                 {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                return S_OK;
            }

            HRESULT STDMETHODCALLTYPE FileSystemItemCompleted(FileSystemOperation /*operationType*/,
                                                              unsigned long /*itemIndex*/,
                                                              const wchar_t* /*sourcePath*/,
                                                              const wchar_t* /*destinationPath*/,
                                                              HRESULT /*status*/,
                                                              FileSystemOptions* /*options*/,
                                                              void* /*cookie*/) noexcept override
            {
                return S_OK;
            }

            HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, void* cookie) noexcept override
            {
                return task.FileSystemShouldCancel(pCancel, cookie);
            }

            HRESULT STDMETHODCALLTYPE FileSystemIssue(FileSystemOperation operationType,
                                                      const wchar_t* sourcePath,
                                                      const wchar_t* destinationPath,
                                                      HRESULT status,
                                                      FileSystemIssueAction* action,
                                                      FileSystemOptions* options,
                                                      void* cookie) noexcept override
            {
                return task.FileSystemIssue(operationType, sourcePath, destinationPath, status, action, options, cookie);
            }
        };

        struct CrossFileSystemBridge
        {
            static constexpr size_t BufferSize() noexcept
            {
                return 1024u * 1024u;
            }
            static constexpr DWORD SleepSliceMs() noexcept
            {
                return 50u;
            }

            Task& task;
            IFileSystem& sourceFs;
            IFileSystem& destinationFs;
            IFileSystemIO& sourceIo;
            IFileSystemIO& destinationIo;
            IFileSystemDirectoryOperations* destinationDirOps = nullptr;
            FileSystemFlags flags                             = FILESYSTEM_FLAG_NONE;
            void* cookie                                      = nullptr;
            DWORD sourceRootAttributesHint                    = 0;
            ReparsePointPolicy reparsePointPolicy             = ReparsePointPolicy::CopyReparse;

            // Total bytes is best-effort: if unknown, keep 0.
            unsigned __int64 totalBytes                 = 0;
            unsigned __int64 completedBytes             = 0;
            unsigned long skippedDirectoryReparseCount  = 0;
            bool rootDirectoryReparseSkipped            = false;
            bool unsupportedDirectoryReparseEncountered = false;

            ULONGLONG startTick = 0;
            FileSystemOptions options{};

            std::unique_ptr<std::byte[]> buffer;
            unsigned long bufferBytes = 0;

            CrossFileSystemBridge(Task& owner,
                                  IFileSystem& source,
                                  IFileSystem& destination,
                                  IFileSystemIO& sourceIoIn,
                                  IFileSystemIO& destinationIoIn,
                                  IFileSystemDirectoryOperations* destinationDirOpsIn,
                                  FileSystemFlags flagsIn,
                                  void* cookieIn,
                                  unsigned __int64 totalBytesIn,
                                  DWORD sourceRootAttributesHintIn,
                                  ReparsePointPolicy reparsePointPolicyIn) noexcept
                : task(owner),
                  sourceFs(source),
                  destinationFs(destination),
                  sourceIo(sourceIoIn),
                  destinationIo(destinationIoIn),
                  destinationDirOps(destinationDirOpsIn),
                  flags(flagsIn),
                  cookie(cookieIn),
                  sourceRootAttributesHint(sourceRootAttributesHintIn),
                  reparsePointPolicy(reparsePointPolicyIn),
                  totalBytes(totalBytesIn)
            {
                options.bandwidthLimitBytesPerSecond = task._desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);

                buffer.reset(new (std::nothrow) std::byte[BufferSize()]);
                bufferBytes = BufferSize() > static_cast<size_t>(std::numeric_limits<unsigned long>::max()) ? std::numeric_limits<unsigned long>::max()
                                                                                                            : static_cast<unsigned long>(BufferSize());
            }

            CrossFileSystemBridge(const CrossFileSystemBridge&)            = delete;
            CrossFileSystemBridge(CrossFileSystemBridge&&)                 = delete;
            CrossFileSystemBridge& operator=(const CrossFileSystemBridge&) = delete;
            CrossFileSystemBridge& operator=(CrossFileSystemBridge&&)      = delete;

            [[nodiscard]] bool CancelRequested() const noexcept
            {
                return task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested();
            }

            void SleepResponsive(DWORD totalMs) noexcept
            {
                while (totalMs > 0)
                {
                    if (CancelRequested())
                    {
                        return;
                    }

                    task.WaitWhilePaused();

                    const DWORD slice = (std::min)(totalMs, SleepSliceMs());
                    ::Sleep(slice);
                    totalMs -= slice;
                }
            }

            void Throttle(unsigned __int64 bytesSoFar) noexcept
            {
                const unsigned __int64 bandwidthLimit = options.bandwidthLimitBytesPerSecond;
                if (bandwidthLimit == 0)
                {
                    return;
                }

                if (startTick == 0)
                {
                    startTick = GetTickCount64();
                }

                const ULONGLONG now              = GetTickCount64();
                const unsigned __int64 elapsedMs = static_cast<unsigned __int64>(now - startTick);

                constexpr unsigned __int64 maxSafeBytes = std::numeric_limits<unsigned __int64>::max() / 1000u;

                unsigned __int64 desiredMs = 0;
                if (bytesSoFar > 0 && bytesSoFar <= maxSafeBytes)
                {
                    desiredMs = (bytesSoFar * 1000u) / bandwidthLimit;
                }
                else if (bytesSoFar > maxSafeBytes)
                {
                    desiredMs = std::numeric_limits<unsigned __int64>::max();
                }

                if (desiredMs > elapsedMs)
                {
                    const unsigned __int64 remaining = desiredMs - elapsedMs;
                    const DWORD sleepMs = remaining > std::numeric_limits<DWORD>::max() ? std::numeric_limits<DWORD>::max() : static_cast<DWORD>(remaining);
                    if (sleepMs > 0)
                    {
                        SleepResponsive(sleepMs);
                    }
                }
            }

            HRESULT ReportProgress(const std::wstring& currentSourcePath,
                                   const std::wstring& currentDestinationPath,
                                   unsigned __int64 currentItemTotalBytes,
                                   unsigned __int64 currentItemCompletedBytes,
                                   unsigned __int64 callCompletedBytes) noexcept
            {
                const unsigned __int64 clampedCallCompleted = (totalBytes > 0) ? (std::min)(totalBytes, callCompletedBytes) : callCompletedBytes;
                return task.FileSystemProgress(task._operation,
                                               1,
                                               0,
                                               totalBytes,
                                               clampedCallCompleted,
                                               currentSourcePath.c_str(),
                                               currentDestinationPath.c_str(),
                                               currentItemTotalBytes,
                                               currentItemCompletedBytes,
                                               &options,
                                               0,
                                               cookie);
            }

            HRESULT EnsureDestinationDirectory(const std::wstring& destinationPath) noexcept
            {
                if (! destinationDirOps)
                {
                    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }

                unsigned long attributes = 0;
                const HRESULT hrAttr     = destinationIo.GetAttributes(destinationPath.c_str(), &attributes);
                if (SUCCEEDED(hrAttr))
                {
                    if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
                    {
                        return S_OK;
                    }

                    if ((flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) == 0)
                    {
                        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                    }

                    // Replace an existing file with a directory.
                    BridgeCallback callback(task);
                    const HRESULT hrDelete = destinationFs.DeleteItem(destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
                    if (FAILED(hrDelete))
                    {
                        return hrDelete;
                    }
                }

                const HRESULT hrCreate = destinationDirOps->CreateDirectory(destinationPath.c_str());
                if (SUCCEEDED(hrCreate) || hrCreate == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
                {
                    return S_OK;
                }

                return hrCreate;
            }

            void MarkDirectoryReparseSkipped(const std::wstring& sourcePath, const std::wstring& destinationPath, bool isRoot) noexcept
            {
                ++skippedDirectoryReparseCount;
                if (isRoot)
                {
                    rootDirectoryReparseSkipped = true;
                }

                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                   HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                                   L"bridge.reparse.skip",
                                   isRoot ? L"Skipped root directory reparse point by policy." : L"Skipped directory reparse point by policy.",
                                   sourcePath,
                                   destinationPath);

                static_cast<void>(ReportProgress(sourcePath, destinationPath, 0, 0, completedBytes));
            }

            HRESULT CopyFile(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
            {
                if (! buffer || bufferBytes == 0)
                {
                    return E_OUTOFMEMORY;
                }

                wil::com_ptr<IFileReader> reader;
                HRESULT hr = sourceIo.CreateFileReader(sourcePath.c_str(), reader.addressof());
                if (FAILED(hr))
                {
                    return hr;
                }

                FileSystemBasicInformation sourceBasicInfo{};
                bool hasSourceBasicInfo  = false;
                const HRESULT hrGetBasic = sourceIo.GetFileBasicInformation(sourcePath.c_str(), &sourceBasicInfo);
                if (SUCCEEDED(hrGetBasic))
                {
                    hasSourceBasicInfo = true;
                }
                else if (hrGetBasic != E_NOTIMPL && hrGetBasic != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
                {
                    Debug::Warning(
                        L"CrossFileSystemBridge: GetFileBasicInformation failed for '{}' (hr={:#x})", sourcePath, static_cast<unsigned long>(hrGetBasic));
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                       hrGetBasic,
                                       L"bridge.metadata.read",
                                       L"GetFileBasicInformation failed for source file.",
                                       sourcePath,
                                       destinationPath);
                }

                unsigned __int64 fileTotalBytes = 0;
                static_cast<void>(reader->GetSize(&fileTotalBytes));
                if (totalBytes == 0 && fileTotalBytes > 0)
                {
                    totalBytes = fileTotalBytes;
                }

                wil::com_ptr<IFileWriter> writer;
                hr = destinationIo.CreateFileWriter(destinationPath.c_str(), flags, writer.addressof());
                if (FAILED(hr))
                {
                    return hr;
                }

                unsigned __int64 fileCompletedBytes = 0;
                hr                                  = ReportProgress(sourcePath, destinationPath, fileTotalBytes, fileCompletedBytes, completedBytes);
                if (FAILED(hr))
                {
                    return hr;
                }

                for (;;)
                {
                    if (CancelRequested())
                    {
                        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }

                    unsigned long bytesRead = 0;
                    hr                      = reader->Read(buffer.get(), bufferBytes, &bytesRead);
                    if (FAILED(hr))
                    {
                        return hr;
                    }

                    if (bytesRead == 0)
                    {
                        break;
                    }

                    size_t offset = 0;
                    while (offset < bytesRead)
                    {
                        if (CancelRequested())
                        {
                            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        }

                        unsigned long bytesWritten  = 0;
                        const unsigned long toWrite = static_cast<unsigned long>(
                            std::min(static_cast<size_t>(bytesRead - offset), static_cast<size_t>(std::numeric_limits<unsigned long>::max())));
                        hr = writer->Write(buffer.get() + offset, toWrite, &bytesWritten);
                        if (FAILED(hr))
                        {
                            return hr;
                        }
                        if (bytesWritten == 0)
                        {
                            return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
                        }

                        offset += bytesWritten;

                        if (fileCompletedBytes > std::numeric_limits<unsigned __int64>::max() - bytesWritten)
                        {
                            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                        }
                        fileCompletedBytes += bytesWritten;

                        const unsigned __int64 callCompleted = completedBytes + fileCompletedBytes;
                        hr                                   = ReportProgress(sourcePath, destinationPath, fileTotalBytes, fileCompletedBytes, callCompleted);
                        if (FAILED(hr))
                        {
                            return hr;
                        }

                        Throttle(callCompleted);
                    }
                }

                if (fileTotalBytes > 0 && fileCompletedBytes >= fileTotalBytes)
                {
                    // Some destination writers (e.g. remote plugins) may perform significant work during Commit()
                    // after the bridge finishes staging writes. For small files this can look like a "stuck at 100%"
                    // progress bar. Switch to an indeterminate item bar during Commit() so the UI stays obviously active.
                    constexpr unsigned __int64 kSmallFileCommitIndeterminateThresholdBytes = 1024ull * 1024ull;
                    if (fileTotalBytes <= kSmallFileCommitIndeterminateThresholdBytes)
                    {
                        const unsigned __int64 callCompleted = (completedBytes > (std::numeric_limits<unsigned __int64>::max)() - fileCompletedBytes)
                                                                  ? (std::numeric_limits<unsigned __int64>::max)()
                                                                  : (completedBytes + fileCompletedBytes);
                        hr = ReportProgress(sourcePath, destinationPath, 0, 0, callCompleted);
                        if (FAILED(hr))
                        {
                            return hr;
                        }
                    }
                }

                hr = writer->Commit();
                if (FAILED(hr))
                {
                    return hr;
                }

                if (hasSourceBasicInfo)
                {
                    const HRESULT hrSetBasic = destinationIo.SetFileBasicInformation(destinationPath.c_str(), &sourceBasicInfo);
                    if (FAILED(hrSetBasic) && hrSetBasic != E_NOTIMPL && hrSetBasic != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
                    {
                        Debug::Warning(L"CrossFileSystemBridge: SetFileBasicInformation failed for '{}' (hr={:#x})",
                                       destinationPath,
                                       static_cast<unsigned long>(hrSetBasic));
                        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                           hrSetBasic,
                                           L"bridge.metadata.write",
                                           L"SetFileBasicInformation failed for destination file.",
                                           sourcePath,
                                           destinationPath);
                    }
                }

                if (completedBytes > std::numeric_limits<unsigned __int64>::max() - fileCompletedBytes)
                {
                    return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                }
                completedBytes += fileCompletedBytes;

                const unsigned __int64 finalTotal     = fileTotalBytes > 0 ? fileTotalBytes : fileCompletedBytes;
                const unsigned __int64 finalCompleted = fileCompletedBytes;

                hr = ReportProgress(sourcePath, destinationPath, finalTotal, finalCompleted, completedBytes);
                if (FAILED(hr))
                {
                    return hr;
                }

                return S_OK;
            }

            HRESULT CopyDirectory(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
            {
                if (CancelRequested())
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                HRESULT hr = EnsureDestinationDirectory(destinationPath);
                if (FAILED(hr))
                {
                    return hr;
                }

                wil::com_ptr<IFilesInformation> info;
                hr = sourceFs.ReadDirectoryInfo(sourcePath.c_str(), info.addressof());
                if (FAILED(hr))
                {
                    return hr;
                }

                FileInfo* entry = nullptr;
                hr              = info->GetBuffer(&entry);
                if (FAILED(hr) || entry == nullptr)
                {
                    return hr;
                }

                unsigned long bufferSize = 0;
                hr                       = info->GetBufferSize(&bufferSize);
                if (FAILED(hr) || bufferSize < sizeof(FileInfo))
                {
                    return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                }

                std::byte* base = reinterpret_cast<std::byte*>(entry);
                std::byte* end  = base + bufferSize;

                for (;;)
                {
                    task.WaitWhilePaused();
                    if (CancelRequested())
                    {
                        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }

                    const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
                    const std::wstring_view name(entry->FileName, nameChars);

                    const bool isDot = (name == L"." || name == L"..");
                    if (! name.empty() && ! isDot)
                    {
                        const bool isDirectory         = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                        const bool isReparse           = (entry->FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                        const std::wstring childSource = JoinFolderAndLeaf(sourcePath, name);
                        const std::wstring childDest   = JoinFolderAndLeaf(destinationPath, name);

                        if (isDirectory)
                        {
                            if (isReparse && reparsePointPolicy != ReparsePointPolicy::FollowTargets)
                            {
                                if (reparsePointPolicy == ReparsePointPolicy::Skip)
                                {
                                    MarkDirectoryReparseSkipped(childSource, childDest, false);
                                    continue;
                                }

                                // copyReparse requires preserving a link; bridge cannot preserve NTFS reparse payloads.
                                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                                   HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                                   L"bridge.reparse.unsupported",
                                                   L"Cross-filesystem bridge cannot preserve directory reparse payloads.",
                                                   childSource,
                                                   childDest);
                                unsupportedDirectoryReparseEncountered = true;
                                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                            }
                            hr = CopyDirectory(childSource, childDest);
                        }
                        else
                        {
                            hr = CopyFile(childSource, childDest);
                        }

                        if (FAILED(hr))
                        {
                            return hr;
                        }
                    }

                    if (entry->NextEntryOffset == 0)
                    {
                        break;
                    }

                    if (entry->NextEntryOffset < sizeof(FileInfo))
                    {
                        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    }

                    std::byte* next = reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset;
                    if (next < base || next + sizeof(FileInfo) > end)
                    {
                        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    }

                    entry = reinterpret_cast<FileInfo*>(next);
                }

                return S_OK;
            }

            HRESULT CopyPath(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
            {
                unsigned long attributes = sourceRootAttributesHint;
                if (attributes == 0)
                {
                    const HRESULT hrAttr = sourceIo.GetAttributes(sourcePath.c_str(), &attributes);
                    if (FAILED(hrAttr))
                    {
                        return hrAttr;
                    }
                }

                const bool isDirectory = (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                const bool isReparse   = (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                if (isDirectory)
                {
                    if (isReparse && reparsePointPolicy != ReparsePointPolicy::FollowTargets)
                    {
                        if (reparsePointPolicy == ReparsePointPolicy::Skip)
                        {
                            MarkDirectoryReparseSkipped(sourcePath, destinationPath, true);
                            return S_OK;
                        }
                        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                           HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                           L"bridge.reparse.unsupported",
                                           L"Cross-filesystem bridge cannot preserve root directory reparse payloads.",
                                           sourcePath,
                                           destinationPath);
                        unsupportedDirectoryReparseEncountered = true;
                        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                    }
                    return CopyDirectory(sourcePath, destinationPath);
                }

                return CopyFile(sourcePath, destinationPath);
            }
        };

        if (_perItemMaxConcurrency > 1u)
        {
            // Per-task multi-item concurrency: run multiple CopyItem/MoveItem/DeleteItem calls concurrently while keeping
            // conflict prompts serialized (one prompt per task at a time).
            std::atomic<size_t> nextIndex{0};
            std::atomic<bool> hadSkipped{false};
            std::atomic<HRESULT> firstFailure{S_OK};

            const auto processIndex = [&](size_t index) noexcept -> HRESULT
            {
                const std::wstring& sourceText = _sourcePaths[index].native();
                if (sourceText.empty())
                {
                    return E_INVALIDARG;
                }

                const unsigned __int64 preCalcBytesForItem = (canUsePreCalcBytes && index < _preCalcSourceBytes.size()) ? _preCalcSourceBytes[index] : 0;

                std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> retryCounts{};
                std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> cachedModifierAttempts{};
                FileSystemFlags itemFlags = _flags;

                bool itemSucceeded                  = false;
                bool itemSkipped                    = false;
                unsigned __int64 callCompletedBytes = 0;
                unsigned __int64 callCompletedItems = 0;
                unsigned __int64 callTotalItems     = 0;

                for (;;)
                {
                    WaitWhilePaused();
                    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
                    {
                        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }

                    std::wstring destinationItemText;
                    if (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE)
                    {
                        const std::wstring_view leaf = GetPathLeaf(sourceText);
                        if (leaf.empty())
                        {
                            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                        }
                        destinationItemText = JoinFolderAndLeaf(destinationFolderText, leaf);
                    }

                    PerItemCallbackCookie cookie{index};

                    {
                        std::scoped_lock lock(_progressMutex);
                        if (_perItemInFlightCallCount < _perItemInFlightCalls.size())
                        {
                            _perItemInFlightCalls[_perItemInFlightCallCount] = {&cookie, 0, 0};
                            ++_perItemInFlightCallCount;
                        }

                        _progressCompletedItems       = (std::max)(_progressCompletedItems, _perItemCompletedItems);
                        const unsigned __int64 mapped = _perItemCompletedBytes + computeInFlightCompletedBytesLocked();
                        _progressCompletedBytes       = (std::max)(_progressCompletedBytes, mapped);
                    }

                    callCompletedBytes = 0;
                    callCompletedItems = 0;
                    callTotalItems     = 0;

                    HRESULT itemHr = E_NOTIMPL;
                    if (_operation == FILESYSTEM_COPY)
                    {
                        FileSystemOptions options{};
                        options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                        itemHr = _fileSystem->CopyItem(sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                    }
                    else if (_operation == FILESYSTEM_MOVE)
                    {
                        FileSystemOptions options{};
                        options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                        itemHr = _fileSystem->MoveItem(sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                    }
                    else if (_operation == FILESYSTEM_DELETE)
                    {
                        itemHr = _fileSystem->DeleteItem(sourceText.c_str(), itemFlags, nullptr, this, static_cast<void*>(&cookie));
                    }

                    {
                        std::scoped_lock lock(_progressMutex);
                        for (size_t i = 0; i < _perItemInFlightCallCount; ++i)
                        {
                            if (_perItemInFlightCalls[i].cookie == &cookie)
                            {
                                callCompletedItems       = _perItemInFlightCalls[i].completedItems;
                                callCompletedBytes       = _perItemInFlightCalls[i].completedBytes;
                                callTotalItems           = static_cast<unsigned __int64>(_perItemInFlightCalls[i].totalItems);
                                _perItemInFlightCalls[i] = _perItemInFlightCalls[_perItemInFlightCallCount - 1u];
                                --_perItemInFlightCallCount;
                                break;
                            }
                        }

                        if (_operation == FILESYSTEM_DELETE)
                        {
                            if (callCompletedItems > 0)
                            {
                                if (_perItemCompletedEntryCount > std::numeric_limits<unsigned __int64>::max() - callCompletedItems)
                                {
                                    _perItemCompletedEntryCount = std::numeric_limits<unsigned __int64>::max();
                                }
                                else
                                {
                                    _perItemCompletedEntryCount += callCompletedItems;
                                }
                            }

                            if (callTotalItems > 0)
                            {
                                if (_perItemTotalEntryCount > std::numeric_limits<unsigned __int64>::max() - callTotalItems)
                                {
                                    _perItemTotalEntryCount = std::numeric_limits<unsigned __int64>::max();
                                }
                                else
                                {
                                    _perItemTotalEntryCount += callTotalItems;
                                }
                            }

                            const unsigned __int64 mappedCompletedItems = _perItemCompletedEntryCount + computeInFlightCompletedItemsLocked();
                            const unsigned __int64 clampedCompleted =
                                std::min<unsigned __int64>(mappedCompletedItems, static_cast<unsigned __int64>(std::numeric_limits<unsigned long>::max()));
                            _progressCompletedItems = (std::max)(_progressCompletedItems, static_cast<unsigned long>(clampedCompleted));

                            const bool precalcTotalAvailable = _preCalcCompleted.load(std::memory_order_acquire) && _progressTotalItems > 0;
                            if (! precalcTotalAvailable)
                            {
                                const unsigned __int64 mappedTotalItems = _perItemTotalEntryCount + computeInFlightTotalItemsLocked();
                                if (mappedTotalItems > 0)
                                {
                                    const unsigned __int64 clampedTotal =
                                        std::min<unsigned __int64>(mappedTotalItems, static_cast<unsigned __int64>(std::numeric_limits<unsigned long>::max()));
                                    _progressTotalItems = (std::max)(_progressTotalItems, static_cast<unsigned long>(clampedTotal));
                                }
                            }
                        }

                        const unsigned __int64 mapped = _perItemCompletedBytes + computeInFlightCompletedBytesLocked();
                        _progressCompletedBytes       = (std::max)(_progressCompletedBytes, mapped);
                    }

                    const bool cancelled = itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || itemHr == E_ABORT;
                    if (cancelled)
                    {
                        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }

                    if (itemHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
                    {
                        itemSucceeded = true;
                        hadSkipped.store(true, std::memory_order_release);
                        break;
                    }

                    if (SUCCEEDED(itemHr))
                    {
                        itemSucceeded = true;
                        break;
                    }

                    if (continueOnError)
                    {
                        auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                        LogDiagnostic(DiagnosticSeverity::Warning,
                                      itemHr,
                                      L"item.continueOnError",
                                      L"Item failed and was skipped due continue-on-error.",
                                      diagnosticSource,
                                      diagnosticDestination);
                        itemSkipped = true;
                        hadSkipped.store(true, std::memory_order_release);
                        break;
                    }

                    const ConflictBucket bucket = ClassifyConflictBucket(_operation, itemFlags, fileSystemIo, itemHr, sourceText, destinationItemText, false);
                    if (bucket == ConflictBucket::RecycleBinFailed)
                    {
                        auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                        LogDiagnostic(DiagnosticSeverity::Error,
                                      itemHr,
                                      L"delete.recycleBin.item",
                                      L"Recycle Bin delete failed for item.",
                                      diagnosticSource,
                                      diagnosticDestination);
                    }

                    const size_t bucketIndex = static_cast<size_t>(bucket);

                    std::optional<ConflictAction> cached = getCachedDecision(bucket);
                    if (cached.has_value() && isModifierConflictAction(cached.value()) && bucketIndex < cachedModifierAttempts.size() &&
                        cachedModifierAttempts[bucketIndex] >= kMaxCachedModifierAttemptsPerBucket)
                    {
                        clearCachedDecision(bucket);
                        cached.reset();
                    }
                    ConflictAction action = cached.value_or(ConflictAction::None);

                    if (action == ConflictAction::None)
                    {
                        const bool canRetryBucket = bucket != ConflictBucket::UnsupportedReparse;
                        const bool allowRetry     = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] == 0u;
                        const bool retryFailed    = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] != 0u;

                        bool owner = false;
                        {
                            std::unique_lock lock(_conflictMutex);

                            const bool cacheableBucket = bucketIndex < _conflictDecisionCache.size();
                            if (cacheableBucket && _conflictDecisionCache[bucketIndex].has_value())
                            {
                                action = _conflictDecisionCache[bucketIndex].value();
                            }
                            else
                            {
                                _conflictCv.wait(
                                    lock,
                                    [&]() noexcept
                                    { return ! _conflictPrompt.active || _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested(); });

                                if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
                                {
                                    action = ConflictAction::Cancel;
                                }
                                else if (cacheableBucket && _conflictDecisionCache[bucketIndex].has_value())
                                {
                                    action = _conflictDecisionCache[bucketIndex].value();
                                }
                                else
                                {
                                    setConflictPromptLocked(&cookie, bucket, itemHr, sourceText, destinationItemText, allowRetry, retryFailed);
                                    owner = true;
                                }
                            }
                        }

                        if (owner)
                        {
                            const auto decision   = waitForConflictDecision();
                            action                = decision.first;
                            const bool applyToAll = decision.second;

                            if (applyToAll && action != ConflictAction::Retry && action != ConflictAction::Cancel && action != ConflictAction::None)
                            {
                                setCachedDecision(bucket, action);
                            }
                        }
                    }

                    if (action == ConflictAction::Overwrite)
                    {
                        if (bucketIndex < cachedModifierAttempts.size())
                        {
                            ++cachedModifierAttempts[bucketIndex];
                        }
                        itemFlags = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
                        continue;
                    }

                    if (action == ConflictAction::ReplaceReadOnly)
                    {
                        if (bucketIndex < cachedModifierAttempts.size())
                        {
                            ++cachedModifierAttempts[bucketIndex];
                        }
                        itemFlags = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
                        continue;
                    }

                    if (action == ConflictAction::PermanentDelete)
                    {
                        if (bucketIndex < cachedModifierAttempts.size())
                        {
                            ++cachedModifierAttempts[bucketIndex];
                        }
                        itemFlags = static_cast<FileSystemFlags>(itemFlags & ~FILESYSTEM_FLAG_USE_RECYCLE_BIN);
                        continue;
                    }

                    if (action == ConflictAction::Retry)
                    {
                        if (bucketIndex < retryCounts.size() && retryCounts[bucketIndex] == 0u)
                        {
                            retryCounts[bucketIndex] = 1u;
                            if (bucket == ConflictBucket::SharingViolation)
                            {
                                Sleep(750);
                            }
                            continue;
                        }
                        action = ConflictAction::Skip;
                    }

                    if (action == ConflictAction::SkipAll)
                    {
                        auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                        LogDiagnostic(DiagnosticSeverity::Warning,
                                      itemHr,
                                      L"item.conflict.skipAll",
                                      L"Conflict action Skip all similar conflicts selected.",
                                      diagnosticSource,
                                      diagnosticDestination);
                        setCachedDecision(bucket, ConflictAction::Skip);
                        itemSkipped = true;
                        hadSkipped.store(true, std::memory_order_release);
                        break;
                    }

                    if (action == ConflictAction::Skip)
                    {
                        auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                        LogDiagnostic(DiagnosticSeverity::Warning,
                                      itemHr,
                                      L"item.conflict.skip",
                                      L"Conflict action Skip item selected.",
                                      diagnosticSource,
                                      diagnosticDestination);
                        itemSkipped = true;
                        hadSkipped.store(true, std::memory_order_release);
                        break;
                    }

                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                if (itemSkipped && preCalcBytesForItem > 0)
                {
                    std::scoped_lock lock(_progressMutex);
                    _progressTotalBytes = (_progressTotalBytes >= preCalcBytesForItem) ? (_progressTotalBytes - preCalcBytesForItem) : 0;
                    // If pre-calc bytes were counted into total, and the user later skips the item,
                    // ensure we don't end up reporting "completed > total" (progress > 100%).
                    _progressCompletedBytes = (std::min)(_progressCompletedBytes, _progressTotalBytes);
                }

                unsigned __int64 bytesForItem = 0;
                if (itemSucceeded)
                {
                    bytesForItem = (preCalcBytesForItem > 0) ? preCalcBytesForItem : callCompletedBytes;
                }

                {
                    std::scoped_lock lock(_progressMutex);
                    if (itemSucceeded)
                    {
                        if (_perItemCompletedBytes > std::numeric_limits<unsigned __int64>::max() - bytesForItem)
                        {
                            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                        }
                        _perItemCompletedBytes += bytesForItem;
                    }

                    if (_perItemCompletedItems < std::numeric_limits<unsigned long>::max())
                    {
                        ++_perItemCompletedItems;
                    }
                    _progressCompletedItems       = (std::max)(_progressCompletedItems, _perItemCompletedItems);
                    const unsigned __int64 mapped = _perItemCompletedBytes + computeInFlightCompletedBytesLocked();
                    _progressCompletedBytes       = (std::max)(_progressCompletedBytes, mapped);
                }

                return S_OK;
            };

            const auto runWorker = [&]() noexcept
            {
                [[maybe_unused]] auto coInit = wil::CoInitializeEx();
                for (;;)
                {
                    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
                    {
                        return;
                    }

                    const size_t index = nextIndex.fetch_add(1u, std::memory_order_acq_rel);
                    if (index >= _sourcePaths.size())
                    {
                        return;
                    }

                    const HRESULT hrItem = processIndex(index);
                    if (FAILED(hrItem))
                    {
                        HRESULT expected = S_OK;
                        firstFailure.compare_exchange_strong(expected, hrItem, std::memory_order_acq_rel);
                        RequestCancel();
                        return;
                    }
                }
            };

            std::vector<std::jthread> workers;
            workers.reserve(_perItemMaxConcurrency - 1u);
            for (unsigned int i = 1u; i < _perItemMaxConcurrency; ++i)
            {
                workers.emplace_back([&](std::stop_token) noexcept { runWorker(); });
            }
            runWorker();

            clearConflictPrompt();

            const HRESULT hr = firstFailure.load(std::memory_order_acquire);
            if (FAILED(hr))
            {
                return hr;
            }

            if (hadSkipped.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }

            return S_OK;
        }

        for (size_t index = 0; index < _sourcePaths.size(); ++index)
        {
            const std::wstring& sourceText = _sourcePaths[index].native();
            if (sourceText.empty())
            {
                return E_INVALIDARG;
            }

            const unsigned __int64 preCalcBytesForItem = (canUsePreCalcBytes && index < _preCalcSourceBytes.size()) ? _preCalcSourceBytes[index] : 0;

            std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> retryCounts{};
            std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> cachedModifierAttempts{};

            bool itemSucceeded        = false;
            bool itemSkipped          = false;
            bool itemPartiallySkipped = false;

            FileSystemFlags itemFlags           = _flags;
            unsigned __int64 callCompletedBytes = 0;
            unsigned __int64 callCompletedItems = 0;
            unsigned __int64 callTotalItems     = 0;
            bool moveCopyCompleted              = false;
            unsigned __int64 moveCopiedBytes    = 0;

            for (;;)
            {
                WaitWhilePaused();
                if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
                {
                    clearConflictPrompt();
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                callCompletedBytes = 0;
                callCompletedItems = 0;
                callTotalItems     = 0;

                std::wstring destinationItemText;
                if (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE)
                {
                    const std::wstring_view leaf = GetPathLeaf(sourceText);
                    if (leaf.empty())
                    {
                        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                    }
                    destinationItemText = JoinFolderAndLeaf(destinationFolderText, leaf);
                }

                PerItemCallbackCookie cookie{index};

                {
                    std::scoped_lock lock(_progressMutex);
                    _perItemCompletedItems =
                        static_cast<unsigned long>(std::min<unsigned __int64>(static_cast<unsigned __int64>(index), static_cast<unsigned __int64>(ULONG_MAX)));
                    _perItemInFlightCallCount = 0;
                    if (! _perItemInFlightCalls.empty())
                    {
                        _perItemInFlightCalls[0]  = {&cookie, 0, 0};
                        _perItemInFlightCallCount = 1;
                    }

                    _progressCompletedItems       = _perItemCompletedItems;
                    const unsigned __int64 mapped = _perItemCompletedBytes + computeInFlightCompletedBytesLocked();
                    _progressCompletedBytes       = (std::max)(_progressCompletedBytes, mapped);
                }

                HRESULT itemHr                                   = E_NOTIMPL;
                bool failedDuringMoveDelete                      = false;
                unsigned long bridgeSkippedDirectoryReparseCount = 0;
                bool bridgeRootDirectoryReparseSkipped           = false;
                bool bridgeUnsupportedDirectoryReparse           = false;

                if (_operation == FILESYSTEM_COPY)
                {
                    if (useCrossFileSystemBridge)
                    {
                        CrossFileSystemBridge bridge(*this,
                                                     *_fileSystem,
                                                     *_destinationFileSystem,
                                                     *fileSystemIo,
                                                     *destinationFileSystemIo,
                                                     destinationDirOps.get(),
                                                     itemFlags,
                                                     static_cast<void*>(&cookie),
                                                     preCalcBytesForItem,
                                                     (index < _sourcePathAttributesHint.size()) ? _sourcePathAttributesHint[index] : 0,
                                                     reparsePointPolicy);
                        itemHr                             = bridge.CopyPath(sourceText, destinationItemText);
                        bridgeSkippedDirectoryReparseCount = bridge.skippedDirectoryReparseCount;
                        bridgeRootDirectoryReparseSkipped  = bridge.rootDirectoryReparseSkipped;
                        bridgeUnsupportedDirectoryReparse  = bridge.unsupportedDirectoryReparseEncountered;
                    }
                    else
                    {
                        FileSystemOptions options{};
                        options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                        itemHr = _fileSystem->CopyItem(sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                    }
                }
                else if (_operation == FILESYSTEM_MOVE)
                {
                    if (useCrossFileSystemBridge)
                    {
                        // For cross-filesystem move: copy + delete. If the copy already succeeded and we're retrying due
                        // to a delete failure, skip re-copying (avoid prompting for overwrite again).
                        if (! moveCopyCompleted)
                        {
                            CrossFileSystemBridge bridge(*this,
                                                         *_fileSystem,
                                                         *_destinationFileSystem,
                                                         *fileSystemIo,
                                                         *destinationFileSystemIo,
                                                         destinationDirOps.get(),
                                                         itemFlags,
                                                         static_cast<void*>(&cookie),
                                                         preCalcBytesForItem,
                                                         (index < _sourcePathAttributesHint.size()) ? _sourcePathAttributesHint[index] : 0,
                                                         reparsePointPolicy);
                            itemHr                             = bridge.CopyPath(sourceText, destinationItemText);
                            bridgeSkippedDirectoryReparseCount = bridge.skippedDirectoryReparseCount;
                            bridgeRootDirectoryReparseSkipped  = bridge.rootDirectoryReparseSkipped;
                            bridgeUnsupportedDirectoryReparse  = bridge.unsupportedDirectoryReparseEncountered;
                            if (SUCCEEDED(itemHr))
                            {
                                moveCopyCompleted = bridgeSkippedDirectoryReparseCount == 0 && ! bridgeRootDirectoryReparseSkipped;
                                moveCopiedBytes   = bridge.completedBytes;
                            }
                        }

                        if (SUCCEEDED(itemHr) && moveCopyCompleted)
                        {
                            // Ensure the in-flight call has the best-known completed-bytes snapshot even when we're only deleting.
                            if (moveCopiedBytes > 0)
                            {
                                FileSystemOptions options{};
                                options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                                 const HRESULT hrProgress             = FileSystemProgress(_operation,
                                                                               1,
                                                                               0,
                                                                               preCalcBytesForItem,
                                                                               moveCopiedBytes,
                                                                               sourceText.c_str(),
                                                                               destinationItemText.c_str(),
                                                                               moveCopiedBytes,
                                                                               moveCopiedBytes,
                                                                               &options,
                                                                               0,
                                                                               static_cast<void*>(&cookie));
                                if (FAILED(hrProgress))
                                {
                                    itemHr = hrProgress;
                                }
                            }
                        }

                        if (SUCCEEDED(itemHr) && moveCopyCompleted)
                        {
                            BridgeCallback callback(*this);
                            itemHr = _fileSystem->DeleteItem(sourceText.c_str(), itemFlags, nullptr, &callback, nullptr);
                            if (FAILED(itemHr))
                            {
                                failedDuringMoveDelete = true;
                            }
                        }
                    }
                    else
                    {
                        FileSystemOptions options{};
                        options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                        itemHr = _fileSystem->MoveItem(sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                    }
                }
                else if (_operation == FILESYSTEM_DELETE)
                {
                    itemHr = _fileSystem->DeleteItem(sourceText.c_str(), itemFlags, nullptr, this, static_cast<void*>(&cookie));
                }

                {
                    std::scoped_lock lock(_progressMutex);
                    for (size_t i = 0; i < _perItemInFlightCallCount; ++i)
                    {
                        if (_perItemInFlightCalls[i].cookie == &cookie)
                        {
                            callCompletedItems       = _perItemInFlightCalls[i].completedItems;
                            callCompletedBytes       = _perItemInFlightCalls[i].completedBytes;
                            callTotalItems           = static_cast<unsigned __int64>(_perItemInFlightCalls[i].totalItems);
                            _perItemInFlightCalls[i] = _perItemInFlightCalls[_perItemInFlightCallCount - 1u];
                            --_perItemInFlightCallCount;
                            break;
                        }
                    }

                    if (_operation == FILESYSTEM_DELETE)
                    {
                        if (callCompletedItems > 0)
                        {
                            if (_perItemCompletedEntryCount > std::numeric_limits<unsigned __int64>::max() - callCompletedItems)
                            {
                                _perItemCompletedEntryCount = std::numeric_limits<unsigned __int64>::max();
                            }
                            else
                            {
                                _perItemCompletedEntryCount += callCompletedItems;
                            }
                        }

                        if (callTotalItems > 0)
                        {
                            if (_perItemTotalEntryCount > std::numeric_limits<unsigned __int64>::max() - callTotalItems)
                            {
                                _perItemTotalEntryCount = std::numeric_limits<unsigned __int64>::max();
                            }
                            else
                            {
                                _perItemTotalEntryCount += callTotalItems;
                            }
                        }

                        const unsigned __int64 mappedCompletedItems = _perItemCompletedEntryCount + computeInFlightCompletedItemsLocked();
                        const unsigned __int64 clampedCompleted =
                            std::min<unsigned __int64>(mappedCompletedItems, static_cast<unsigned __int64>(std::numeric_limits<unsigned long>::max()));
                        _progressCompletedItems = (std::max)(_progressCompletedItems, static_cast<unsigned long>(clampedCompleted));

                        const bool precalcTotalAvailable = _preCalcCompleted.load(std::memory_order_acquire) && _progressTotalItems > 0;
                        if (! precalcTotalAvailable)
                        {
                            const unsigned __int64 mappedTotalItems = _perItemTotalEntryCount + computeInFlightTotalItemsLocked();
                            if (mappedTotalItems > 0)
                            {
                                const unsigned __int64 clampedTotal =
                                    std::min<unsigned __int64>(mappedTotalItems, static_cast<unsigned __int64>(std::numeric_limits<unsigned long>::max()));
                                _progressTotalItems = (std::max)(_progressTotalItems, static_cast<unsigned long>(clampedTotal));
                            }
                        }
                    }

                    const unsigned __int64 mapped = _perItemCompletedBytes + computeInFlightCompletedBytesLocked();
                    _progressCompletedBytes       = (std::max)(_progressCompletedBytes, mapped);
                }

                const bool cancelled = itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || itemHr == E_ABORT;
                if (cancelled)
                {
                    clearConflictPrompt();
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                if (itemHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
                {
                    itemPartiallySkipped = true;
                    hadSkippedItems      = true;
                    itemSucceeded        = true;
                    break;
                }

                if (SUCCEEDED(itemHr))
                {
                    if (useCrossFileSystemBridge && bridgeRootDirectoryReparseSkipped)
                    {
                        LogDiagnostic(DiagnosticSeverity::Warning,
                                      HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                                      L"bridge.reparse.skip",
                                      L"Skipped root directory reparse point during bridge operation.",
                                      sourceText,
                                      destinationItemText);
                        itemSkipped     = true;
                        hadSkippedItems = true;
                        break;
                    }

                    if (useCrossFileSystemBridge && bridgeSkippedDirectoryReparseCount > 0)
                    {
                        const std::wstring skipMessage = std::format(L"Skipped {:L} directory reparse point{:s} during bridge operation.",
                                                                     bridgeSkippedDirectoryReparseCount,
                                                                     bridgeSkippedDirectoryReparseCount == 1ul ? L"" : L"s");
                        LogDiagnostic(DiagnosticSeverity::Warning,
                                      HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                                      L"bridge.reparse.skip",
                                      skipMessage,
                                      sourceText,
                                      destinationItemText);
                        itemPartiallySkipped = true;
                        hadSkippedItems      = true;
                    }

                    itemSucceeded = true;
                    break;
                }

                // If the caller explicitly requested continue-on-error, preserve legacy behavior.
                if (continueOnError)
                {
                    auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                    LogDiagnostic(DiagnosticSeverity::Warning,
                                  itemHr,
                                  L"item.continueOnError",
                                  L"Item failed and was skipped due continue-on-error.",
                                  diagnosticSource,
                                  diagnosticDestination);
                    itemSkipped     = true;
                    hadSkippedItems = true;
                    break;
                }

                const FileSystemOperation bucketOperation = failedDuringMoveDelete ? FILESYSTEM_DELETE : _operation;
                const wil::com_ptr<IFileSystemIO>& bucketFileSystemIo =
                    failedDuringMoveDelete ? fileSystemIo : (useCrossFileSystemBridge ? destinationFileSystemIo : fileSystemIo);
                const bool unsupportedReparseHint = bridgeUnsupportedDirectoryReparse;

                const ConflictBucket bucket =
                    ClassifyConflictBucket(bucketOperation, itemFlags, bucketFileSystemIo, itemHr, sourceText, destinationItemText, unsupportedReparseHint);
                if (bucket == ConflictBucket::RecycleBinFailed)
                {
                    auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                    LogDiagnostic(DiagnosticSeverity::Error,
                                  itemHr,
                                  L"delete.recycleBin.item",
                                  L"Recycle Bin delete failed for item.",
                                  diagnosticSource,
                                  diagnosticDestination);
                }

                const size_t bucketIndex = static_cast<size_t>(bucket);

                std::optional<ConflictAction> cached = getCachedDecision(bucket);
                if (cached.has_value() && isModifierConflictAction(cached.value()) && bucketIndex < cachedModifierAttempts.size() &&
                    cachedModifierAttempts[bucketIndex] >= kMaxCachedModifierAttemptsPerBucket)
                {
                    clearCachedDecision(bucket);
                    cached.reset();
                }
                ConflictAction action = cached.value_or(ConflictAction::None);

                if (action == ConflictAction::None)
                {
                    const bool canRetryBucket = bucket != ConflictBucket::UnsupportedReparse;
                    const bool allowRetry     = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] == 0u;
                    const bool retryFailed    = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] != 0u;

                    {
                        std::unique_lock lock(_conflictMutex);
                        setConflictPromptLocked(&cookie, bucket, itemHr, sourceText, destinationItemText, allowRetry, retryFailed);
                    }
                    const auto decision   = waitForConflictDecision();
                    action                = decision.first;
                    const bool applyToAll = decision.second;

                    if (applyToAll && action != ConflictAction::Retry && action != ConflictAction::Cancel && action != ConflictAction::None)
                    {
                        setCachedDecision(bucket, action);
                    }
                }

                if (action == ConflictAction::Overwrite)
                {
                    if (bucketIndex < cachedModifierAttempts.size())
                    {
                        ++cachedModifierAttempts[bucketIndex];
                    }
                    itemFlags = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
                    continue;
                }

                if (action == ConflictAction::ReplaceReadOnly)
                {
                    if (bucketIndex < cachedModifierAttempts.size())
                    {
                        ++cachedModifierAttempts[bucketIndex];
                    }
                    itemFlags = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
                    continue;
                }

                if (action == ConflictAction::PermanentDelete)
                {
                    if (bucketIndex < cachedModifierAttempts.size())
                    {
                        ++cachedModifierAttempts[bucketIndex];
                    }
                    itemFlags = static_cast<FileSystemFlags>(itemFlags & ~FILESYSTEM_FLAG_USE_RECYCLE_BIN);
                    continue;
                }

                if (action == ConflictAction::Retry)
                {
                    if (bucketIndex < retryCounts.size() && retryCounts[bucketIndex] == 0u)
                    {
                        retryCounts[bucketIndex] = 1u;

                        if (bucket == ConflictBucket::SharingViolation)
                        {
                            Sleep(750);
                        }

                        continue;
                    }

                    action = ConflictAction::Skip;
                }

                if (action == ConflictAction::SkipAll)
                {
                    auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                    LogDiagnostic(DiagnosticSeverity::Warning,
                                  itemHr,
                                  L"item.conflict.skipAll",
                                  L"Conflict action Skip all similar conflicts selected.",
                                  diagnosticSource,
                                  diagnosticDestination);
                    setCachedDecision(bucket, ConflictAction::Skip);
                    itemSkipped     = true;
                    hadSkippedItems = true;
                    break;
                }

                if (action == ConflictAction::Skip)
                {
                    auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                    LogDiagnostic(DiagnosticSeverity::Warning,
                                  itemHr,
                                  L"item.conflict.skip",
                                  L"Conflict action Skip item selected.",
                                  diagnosticSource,
                                  diagnosticDestination);
                    itemSkipped     = true;
                    hadSkippedItems = true;
                    break;
                }

                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            if (itemSkipped)
            {
                if (preCalcBytesForItem > 0)
                {
                    std::scoped_lock lock(_progressMutex);
                    _progressTotalBytes = (_progressTotalBytes >= preCalcBytesForItem) ? (_progressTotalBytes - preCalcBytesForItem) : 0;
                    // If pre-calc bytes were counted into total, and the user later skips the item,
                    // ensure we don't end up reporting "completed > total" (progress > 100%).
                    _progressCompletedBytes = (std::min)(_progressCompletedBytes, _progressTotalBytes);
                }
            }
            else if (itemSucceeded || itemPartiallySkipped)
            {
                const unsigned __int64 bytesForItem = (preCalcBytesForItem > 0) ? preCalcBytesForItem : callCompletedBytes;
                if (_perItemCompletedBytes > std::numeric_limits<unsigned __int64>::max() - bytesForItem)
                {
                    clearConflictPrompt();
                    return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                }

                _perItemCompletedBytes += bytesForItem;
            }

            _perItemCompletedItems =
                static_cast<unsigned long>(std::min<unsigned __int64>(static_cast<unsigned __int64>(index + 1u), static_cast<unsigned __int64>(ULONG_MAX)));

            {
                std::scoped_lock lock(_progressMutex);
                _progressCompletedItems       = _perItemCompletedItems;
                const unsigned __int64 mapped = _perItemCompletedBytes + computeInFlightCompletedBytesLocked();
                _progressCompletedBytes       = (std::max)(_progressCompletedBytes, mapped);
            }
        }

        clearConflictPrompt();

        if (hadSkippedItems)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return S_OK;
    }

    if ((_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE) && _destinationFileSystem)
    {
        // Cross-filesystem bridge is only implemented in per-item mode.
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** pathArray = nullptr;
    unsigned long count       = 0;
    HRESULT hr                = BuildPathArrayArena(_sourcePaths, arenaOwner, &pathArray, &count);
    if (FAILED(hr))
    {
        return hr;
    }

    if (count == 0)
    {
        return S_FALSE;
    }

    if (_operation == FILESYSTEM_COPY)
    {
        if (destinationFolder.empty())
        {
            return E_INVALIDARG;
        }

        FileSystemOptions options{};
        options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        return _fileSystem->CopyItems(pathArray, count, destinationFolder.c_str(), _flags, &options, this, nullptr);
    }

    if (_operation == FILESYSTEM_MOVE)
    {
        if (destinationFolder.empty())
        {
            return E_INVALIDARG;
        }

        FileSystemOptions options{};
        options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        return _fileSystem->MoveItems(pathArray, count, destinationFolder.c_str(), _flags, &options, this, nullptr);
    }

    if (_operation == FILESYSTEM_DELETE)
    {
        return _fileSystem->DeleteItems(pathArray, count, _flags, nullptr, this, nullptr);
    }

    return E_NOTIMPL;
}

void FolderWindow::FileOperationState::Task::LogDiagnostic(DiagnosticSeverity severity,
                                                           HRESULT status,
                                                           std::wstring_view category,
                                                           std::wstring_view message,
                                                           std::wstring_view sourcePath,
                                                           std::wstring_view destinationPath) noexcept
{
    if (! _state)
    {
        return;
    }

    std::wstring effectiveSource;
    std::wstring effectiveDestination;

    if (sourcePath.empty() || destinationPath.empty())
    {
        std::scoped_lock lock(_progressMutex);
        if (sourcePath.empty())
        {
            effectiveSource = _progressSourcePath;
        }
        if (destinationPath.empty())
        {
            effectiveDestination = _progressDestinationPath;
        }
    }

    if (! sourcePath.empty())
    {
        effectiveSource = std::wstring(sourcePath);
    }
    if (! destinationPath.empty())
    {
        effectiveDestination = std::wstring(destinationPath);
    }

    _state->RecordTaskDiagnostic(_taskId, _operation, severity, status, category, message, effectiveSource, effectiveDestination);
}

HRESULT FolderWindow::FileOperationState::Task::BuildPathArrayArena(const std::vector<std::filesystem::path>& paths,
                                                                    FileSystemArenaOwner& arenaOwner,
                                                                    const wchar_t*** outPaths,
                                                                    unsigned long* outCount) noexcept
{
    if (! outPaths || ! outCount)
    {
        return E_POINTER;
    }

    *outPaths = nullptr;
    *outCount = 0;

    if (paths.empty())
    {
        return S_OK;
    }

    const unsigned __int64 count64 = static_cast<unsigned __int64>(paths.size());
    if (count64 > std::numeric_limits<unsigned long>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const unsigned __int64 arrayBytes64 = count64 * static_cast<unsigned __int64>(sizeof(const wchar_t*));
    if (arrayBytes64 > std::numeric_limits<unsigned long>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    unsigned long totalBytes = static_cast<unsigned long>(arrayBytes64);

    for (const auto& path : paths)
    {
        const std::wstring& text = path.native();
        const size_t length      = text.size();
        if (length > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)) - 1u)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        const unsigned long bytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
        if (totalBytes > std::numeric_limits<unsigned long>::max() - bytes)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        totalBytes += bytes;
    }

    HRESULT hr = arenaOwner.Initialize(totalBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    FileSystemArena* arena = arenaOwner.Get();
    auto* array            = static_cast<const wchar_t**>(
        AllocateFromFileSystemArena(arena, static_cast<unsigned long>(arrayBytes64), static_cast<unsigned long>(alignof(const wchar_t*))));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    for (size_t index = 0; index < paths.size(); ++index)
    {
        const std::wstring& text  = paths[index].native();
        const size_t length       = text.size();
        const unsigned long bytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
        auto* buffer              = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, bytes, static_cast<unsigned long>(alignof(wchar_t))));
        if (! buffer)
        {
            return E_OUTOFMEMORY;
        }

        if (length > 0)
        {
            ::CopyMemory(buffer, text.data(), length * sizeof(wchar_t));
        }
        buffer[length] = L'\0';
        array[index]   = buffer;
    }

    *outPaths = array;
    *outCount = static_cast<unsigned long>(count64);
    return S_OK;
}

FolderWindow::FileOperationState::FileOperationState(FolderWindow& owner) : _owner(owner)
{
}

FolderWindow::FileOperationState::~FileOperationState()
{
    Shutdown();
}

HRESULT FolderWindow::FileOperationState::StartOperation(FileSystemOperation operation,
                                                         FolderWindow::Pane sourcePane,
                                                         std::optional<FolderWindow::Pane> destinationPane,
                                                         const wil::com_ptr<IFileSystem>& fileSystem,
                                                         std::vector<std::filesystem::path> sourcePaths,
                                                         std::filesystem::path destinationFolder,
                                                         FileSystemFlags flags,
                                                         bool waitForOthers,
                                                         unsigned __int64 initialSpeedLimitBytesPerSecond,
                                                         ExecutionMode executionMode,
                                                         bool requireConfirmation,
                                                         wil::com_ptr<IFileSystem> destinationFileSystem)
{
    if (! fileSystem)
    {
        Debug::Error(L"FolderWindow StartOperation null filesystem");
        return E_POINTER;
    }

    if (sourcePaths.empty())
    {
        Debug::Error(L"FolderWindow StartOperation sourcePath empty");
        return S_FALSE;
    }

    const std::wstring& sourcePluginId = sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.pluginId : _owner._rightPane.pluginId;
    const std::wstring& sourcePluginShortId =
        sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.pluginShortId : _owner._rightPane.pluginShortId;

    const bool allowPreCalcForOperation = operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE ||
                                          // For Recycle Bin deletes, the shell can provide progress without blocking on a full recursive preflight scan.
                                          (operation == FILESYSTEM_DELETE &&
                                           ((flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) == 0 || ! NavigationLocation::IsFilePluginShortId(sourcePluginShortId)));
    const bool enablePreCalc = allowPreCalcForOperation;

    std::vector<DWORD> sourcePathAttributesHint;

    if (operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE)
    {
        unsigned long long fileCount    = 0;
        unsigned long long folderCount  = 0;
        unsigned long long unknownCount = 0;
        std::filesystem::path sampleFile;
        bool hasSampleFile = false;

        const FolderView& sourceFolderView = sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.folderView : _owner._rightPane.folderView;
        const std::vector<FolderView::PathAttributes> selected = sourceFolderView.GetSelectedOrFocusedPathAttributes();
        bool selectionMatches                                  = ! selected.empty() && selected.size() == sourcePaths.size();
        if (selectionMatches)
        {
            for (size_t i = 0; i < selected.size(); ++i)
            {
                if (selected[i].path != sourcePaths[i])
                {
                    selectionMatches = false;
                    break;
                }
            }
        }

        if (selectionMatches)
        {
            for (const auto& item : selected)
            {
                const bool isDirectory = (item.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (isDirectory)
                {
                    ++folderCount;
                    continue;
                }

                ++fileCount;
                if (! hasSampleFile)
                {
                    sampleFile    = item.path;
                    hasSampleFile = true;
                }
            }

            sourcePathAttributesHint.reserve(selected.size());
            for (const auto& item : selected)
            {
                sourcePathAttributesHint.push_back(item.fileAttributes);
            }
        }
        else
        {
            unknownCount = static_cast<unsigned long long>(sourcePaths.size());
        }

        auto suffixFor = [](unsigned long long count) noexcept -> std::wstring_view
        { return count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s"); };

        const unsigned long long itemCount = static_cast<unsigned long long>(sourcePaths.size());
        std::wstring what;
        if (unknownCount > 0)
        {
            const std::wstring_view itemSuffix = suffixFor(itemCount);
            what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_ITEM, itemCount, itemSuffix);
        }
        else if (fileCount > 0 && folderCount > 0)
        {
            const std::wstring_view fileSuffix   = suffixFor(fileCount);
            const std::wstring_view folderSuffix = suffixFor(folderCount);
            what = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILES_FOLDERS, fileCount, fileSuffix, folderCount, folderSuffix);
        }
        else if (fileCount > 0)
        {
            const std::wstring_view fileSuffix = suffixFor(fileCount);
            what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILE, fileCount, fileSuffix);
        }
        else
        {
            const std::wstring_view folderSuffix = suffixFor(folderCount);
            what                                 = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FOLDER, folderCount, folderSuffix);
        }

        auto ensureTrailingSeparator = [](std::wstring text) noexcept -> std::wstring
        {
            if (text.empty())
            {
                return text;
            }

            const wchar_t last = text.back();
            if (last == L'\\' || last == L'/')
            {
                return text;
            }

            text.push_back(L'\\');
            return text;
        };

        auto normalizeSlashes = [](std::wstring& text) noexcept
        {
            for (auto& ch : text)
            {
                if (ch == L'/')
                {
                    ch = L'\\';
                }
            }
        };

        std::wstring fromText;
        if (sourcePaths.size() == 1u)
        {
            fromText = sourcePaths.front().wstring();
            if (unknownCount == 0 && folderCount == 1ull && fileCount == 0ull)
            {
                fromText = ensureTrailingSeparator(std::move(fromText));
            }
        }
        else
        {
            std::filesystem::path commonParent = sourcePaths.front().parent_path();
            bool multipleParents               = false;
            for (size_t index = 1; index < sourcePaths.size(); ++index)
            {
                const std::filesystem::path parent = sourcePaths[index].parent_path();
                if (CompareStringOrdinal(commonParent.c_str(), -1, parent.c_str(), -1, TRUE) != CSTR_EQUAL)
                {
                    multipleParents = true;
                    break;
                }
            }

            if (multipleParents)
            {
                fromText = LoadStringResource(nullptr, IDS_FILEOPS_LOCATION_MULTIPLE);
            }
            else if (unknownCount == 0 && fileCount > 0 && folderCount > 0 && hasSampleFile)
            {
                fromText = sampleFile.wstring();
            }
            else
            {
                fromText = ensureTrailingSeparator(commonParent.wstring());
            }
        }

        std::wstring toText;
        toText = ensureTrailingSeparator(destinationFolder.wstring());
        normalizeSlashes(fromText);
        normalizeSlashes(toText);

        const UINT messageId = operation == FILESYSTEM_COPY ? static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_COPY) : static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_MOVE);
        const std::wstring message = FormatStringResource(nullptr, messageId, what, fromText, toText);

        const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_CONFIRM);

        HostPromptRequest prompt{};
        prompt.version       = 1;
        prompt.sizeBytes     = sizeof(prompt);
        prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
        prompt.severity      = HOST_ALERT_INFO;
        prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
        prompt.targetWindow  = _owner.GetHwnd();
        prompt.title         = caption.c_str();
        prompt.message       = message.c_str();
        prompt.defaultResult = HOST_PROMPT_RESULT_OK;

        HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
        const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
        if (FAILED(hrPrompt) || promptResult != HOST_PROMPT_RESULT_OK)
        {
            return S_FALSE;
        }

        const bool isRecursive = (flags & FILESYSTEM_FLAG_RECURSIVE) != 0;
        if (isRecursive && _owner._settings && GetReparsePointPolicyFromSettings(*_owner._settings, sourcePluginId) == ReparsePointPolicy::FollowTargets)
        {
            bool shouldPrompt = false;
            {
                std::scoped_lock lock(_followTargetsWarningMutex);
                if (_followTargetsWarningAccepted)
                {
                    shouldPrompt = false;
                }
                else if (_followTargetsWarningPromptActive)
                {
                    // Safety-first: if a warning prompt is already visible (possible re-entrancy), abort this start.
                    return S_FALSE;
                }
                else
                {
                    _followTargetsWarningPromptActive = true;
                    shouldPrompt                      = true;
                }
            }

            if (shouldPrompt)
            {
                const std::wstring warningCaption = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
                const std::wstring warningMessage = LoadStringResource(nullptr, IDS_MSG_FILEOPS_REPARSE_FOLLOW_WARNING);

                HostPromptRequest warningPrompt{};
                warningPrompt.version       = 1;
                warningPrompt.sizeBytes     = sizeof(warningPrompt);
                warningPrompt.scope         = HOST_ALERT_SCOPE_WINDOW;
                warningPrompt.severity      = HOST_ALERT_WARNING;
                warningPrompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
                warningPrompt.targetWindow  = _owner.GetHwnd();
                warningPrompt.title         = warningCaption.c_str();
                warningPrompt.message       = warningMessage.c_str();
                warningPrompt.defaultResult = HOST_PROMPT_RESULT_CANCEL;

                HostPromptResult warningResult = HOST_PROMPT_RESULT_NONE;
                const HRESULT hrWarning        = HostShowPrompt(warningPrompt, nullptr, &warningResult);
                if (FAILED(hrWarning) || warningResult != HOST_PROMPT_RESULT_OK)
                {
                    std::scoped_lock lock(_followTargetsWarningMutex);
                    _followTargetsWarningPromptActive = false;
                    return S_FALSE;
                }

                std::scoped_lock lock(_followTargetsWarningMutex);
                _followTargetsWarningPromptActive = false;
                _followTargetsWarningAccepted     = true;
            }
        }
    }
    else if (operation == FILESYSTEM_DELETE && requireConfirmation)
    {
        unsigned long long fileCount    = 0;
        unsigned long long folderCount  = 0;
        unsigned long long unknownCount = 0;
        std::filesystem::path sampleFile;
        bool hasSampleFile = false;

        const FolderView& sourceFolderView = sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.folderView : _owner._rightPane.folderView;
        const std::vector<FolderView::PathAttributes> selected = sourceFolderView.GetSelectedOrFocusedPathAttributes();
        bool selectionMatches                                  = ! selected.empty() && selected.size() == sourcePaths.size();
        if (selectionMatches)
        {
            for (size_t i = 0; i < selected.size(); ++i)
            {
                if (selected[i].path != sourcePaths[i])
                {
                    selectionMatches = false;
                    break;
                }
            }
        }

        if (selectionMatches)
        {
            for (const auto& item : selected)
            {
                const bool isDirectory = (item.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (isDirectory)
                {
                    ++folderCount;
                    continue;
                }

                ++fileCount;
                if (! hasSampleFile)
                {
                    sampleFile    = item.path;
                    hasSampleFile = true;
                }
            }
        }
        else
        {
            unknownCount = static_cast<unsigned long long>(sourcePaths.size());
        }

        auto suffixFor = [](unsigned long long count) noexcept -> std::wstring_view
        { return count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s"); };

        const unsigned long long itemCount = static_cast<unsigned long long>(sourcePaths.size());
        std::wstring what;
        if (unknownCount > 0)
        {
            const std::wstring_view itemSuffix = suffixFor(itemCount);
            what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_ITEM, itemCount, itemSuffix);
        }
        else if (fileCount > 0 && folderCount > 0)
        {
            const std::wstring_view fileSuffix   = suffixFor(fileCount);
            const std::wstring_view folderSuffix = suffixFor(folderCount);
            what = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILES_FOLDERS, fileCount, fileSuffix, folderCount, folderSuffix);
        }
        else if (fileCount > 0)
        {
            const std::wstring_view fileSuffix = suffixFor(fileCount);
            what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILE, fileCount, fileSuffix);
        }
        else
        {
            const std::wstring_view folderSuffix = suffixFor(folderCount);
            what                                 = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FOLDER, folderCount, folderSuffix);
        }

        auto ensureTrailingSeparator = [](std::wstring text) noexcept -> std::wstring
        {
            if (text.empty())
            {
                return text;
            }

            const wchar_t last = text.back();
            if (last == L'\\' || last == L'/')
            {
                return text;
            }

            text.push_back(L'\\');
            return text;
        };

        auto normalizeSlashes = [](std::wstring& text) noexcept
        {
            for (auto& ch : text)
            {
                if (ch == L'/')
                {
                    ch = L'\\';
                }
            }
        };

        std::wstring fromText;
        if (sourcePaths.size() == 1u)
        {
            fromText = sourcePaths.front().wstring();
            if (unknownCount == 0 && folderCount == 1ull && fileCount == 0ull)
            {
                fromText = ensureTrailingSeparator(std::move(fromText));
            }
        }
        else
        {
            std::filesystem::path commonParent = sourcePaths.front().parent_path();
            bool multipleParents               = false;
            for (size_t index = 1; index < sourcePaths.size(); ++index)
            {
                const std::filesystem::path parent = sourcePaths[index].parent_path();
                if (CompareStringOrdinal(commonParent.c_str(), -1, parent.c_str(), -1, TRUE) != CSTR_EQUAL)
                {
                    multipleParents = true;
                    break;
                }
            }

            if (multipleParents)
            {
                fromText = LoadStringResource(nullptr, IDS_FILEOPS_LOCATION_MULTIPLE);
            }
            else if (unknownCount == 0 && fileCount > 0 && folderCount > 0 && hasSampleFile)
            {
                fromText = sampleFile.wstring();
            }
            else
            {
                fromText = ensureTrailingSeparator(commonParent.wstring());
            }
        }

        normalizeSlashes(fromText);

        const std::wstring message = FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFIRM_PERMANENT_DELETE, what, fromText);
        const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_CONFIRM);

        HostPromptRequest prompt{};
        prompt.version       = 1;
        prompt.sizeBytes     = sizeof(prompt);
        prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
        prompt.severity      = HOST_ALERT_WARNING;
        prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
        prompt.targetWindow  = _owner.GetHwnd();
        prompt.title         = caption.c_str();
        prompt.message       = message.c_str();
        prompt.defaultResult = HOST_PROMPT_RESULT_CANCEL;

        HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
        const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
        if (FAILED(hrPrompt) || promptResult != HOST_PROMPT_RESULT_OK)
        {
            return S_FALSE;
        }
    }

    if (operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE)
    {
        auto normalizeSlashes = [](std::wstring& text) noexcept
        {
            for (auto& ch : text)
            {
                if (ch == L'/')
                {
                    ch = L'\\';
                }
            }
        };

        const std::wstring destinationFolderText = destinationFolder.wstring();

        const bool haveAttributesHint = sourcePathAttributesHint.size() == sourcePaths.size();

        std::wstring invalidSourceText;
        std::wstring invalidDestinationItemText;
        for (size_t index = 0; index < sourcePaths.size(); ++index)
        {
            const bool hintIsDirectory = haveAttributesHint && ((sourcePathAttributesHint[index] & FILE_ATTRIBUTE_DIRECTORY) != 0);
            // If we have hints, only directories can cause "copy into self/descendant" recursion.
            // If we don't have hints, be conservative and validate all sources.
            if (haveAttributesHint && ! hintIsDirectory)
            {
                continue;
            }

            const std::wstring sourceText = sourcePaths[index].wstring();
            const std::wstring_view leaf  = GetPathLeaf(sourceText);
            if (leaf.empty())
            {
                continue;
            }

            const std::wstring destinationItemText = JoinFolderAndLeaf(destinationFolderText, leaf);

            std::wstring sourceNormalized          = sourceText;
            std::wstring destinationItemNormalized = destinationItemText;
            normalizeSlashes(sourceNormalized);
            normalizeSlashes(destinationItemNormalized);

            if (! IsSameOrChildPath(sourceNormalized, destinationItemNormalized))
            {
                continue;
            }

            invalidSourceText          = sourceText;
            invalidDestinationItemText = destinationItemText;
            break;
        }

        if (! invalidSourceText.empty())
        {
            Debug::Error(L"FolderWindow StartOperation rejected overlapping destination op={} src:{} dstFolder:{} dstItem:{}",
                         OperationToString(operation),
                         invalidSourceText,
                         destinationFolder.native(),
                         invalidDestinationItemText);

            const std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
            const std::wstring message =
                FormatStringResource(nullptr, IDS_FMT_FILEOPS_INVALID_DESTINATION_OVERLAP, invalidSourceText, destinationFolder.native());
            FolderView& view = sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.folderView : _owner._rightPane.folderView;
            view.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, title, message);
            return S_FALSE;
        }
    }

    auto task = std::make_unique<Task>(*this);
    {
        std::scoped_lock lock(_mutex);
        task->_taskId                   = _nextTaskId++;
        task->_operation                = operation;
        task->_executionMode            = executionMode;
        task->_sourcePane               = sourcePane;
        task->_destinationPane          = destinationPane;
        task->_fileSystem               = fileSystem;
        task->_destinationFileSystem    = std::move(destinationFileSystem);
        task->_sourcePaths              = std::move(sourcePaths);
        task->_sourcePathAttributesHint = std::move(sourcePathAttributesHint);
        task->_destinationFolder        = std::move(destinationFolder);
        task->_flags                    = flags;
        task->_enablePreCalc            = enablePreCalc;
        task->_waitForOthers.store(waitForOthers, std::memory_order_release);
        task->_desiredSpeedLimitBytesPerSecond.store(initialSpeedLimitBytesPerSecond, std::memory_order_release);
        // Mark as waiting in queue immediately if queuing, so UI shows "Waiting..." right away
        task->_waitingInQueue.store(waitForOthers, std::memory_order_release);
    }

    {
        std::scoped_lock lock(task->_progressMutex);
        if (! task->_sourcePaths.empty())
        {
            task->_progressSourcePath = task->_sourcePaths.front().native();
        }

        if (! task->_destinationFolder.empty())
        {
            task->_progressDestinationPath = task->_destinationFolder.native();
        }
    }

    Task* rawTask = task.get();

    {
        std::scoped_lock lock(_mutex);
        _tasks.emplace_back(std::move(task));
    }

    CreateProgressDialog(*rawTask);

    rawTask->_thread = std::jthread([rawTask](std::stop_token stopToken) noexcept { rawTask->ThreadMain(stopToken); });
    return S_OK;
}

void FolderWindow::FileOperationState::ApplyTheme(const AppTheme& /*theme*/)
{
    HWND popup      = nullptr;
    HWND issuesPane = nullptr;
    {
        std::scoped_lock lock(_mutex);
        popup      = _popup.get();
        issuesPane = _issuesPane.get();
    }

    if (popup)
    {
        PostMessageW(popup, WM_THEMECHANGED, 0, 0);
    }

    if (issuesPane)
    {
        PostMessageW(issuesPane, WM_THEMECHANGED, 0, 0);
    }
}

void FolderWindow::FileOperationState::Shutdown() noexcept
{
    std::vector<std::unique_ptr<Task>> tasks;
    wil::unique_hwnd popupToClose;
    wil::unique_hwnd issuesPaneToClose;
    {
        std::scoped_lock lock(_mutex);
        tasks.swap(_tasks);
        popupToClose      = std::move(_popup);
        issuesPaneToClose = std::move(_issuesPane);
    }

    for (auto& task : tasks)
    {
        if (task)
        {
            task->RequestCancel();
        }
    }

    // tasks destruct here; jthread joins automatically.
    FlushDiagnostics(true);
}

void FolderWindow::FileOperationState::NotifyQueueChanged()
{
    _queueCv.notify_all();
}

bool FolderWindow::FileOperationState::HasActiveOperations() noexcept
{
    {
        std::scoped_lock lock(_mutex);
        if (! _tasks.empty())
        {
            return true;
        }
    }

    // Defensive fallback: active operations are expected to always have a task object.
    std::scoped_lock lock(_queueMutex);
    return _activeOperations > 0 || ! _queue.empty();
}

bool FolderWindow::FileOperationState::ShouldQueueNewTask() noexcept
{
    if (! _queueNewTasks.load(std::memory_order_acquire))
    {
        return false;
    }

    return HasActiveOperations();
}

void FolderWindow::FileOperationState::SetQueueNewTasks(bool queue) noexcept
{
    _queueNewTasks.store(queue, std::memory_order_release);
}

bool FolderWindow::FileOperationState::GetQueueNewTasks() const noexcept
{
    return _queueNewTasks.load(std::memory_order_acquire);
}

void FolderWindow::FileOperationState::ApplyQueueMode(bool queue) noexcept
{
    _queueNewTasks.store(queue, std::memory_order_release);

    std::vector<Task*> tasks;
    CollectTasks(tasks);

    for (auto* task : tasks)
    {
        if (! task)
        {
            continue;
        }

        if (! queue)
        {
            task->SetWaitForOthers(false);
            continue;
        }

        if (! task->HasStarted())
        {
            task->SetWaitForOthers(true);
            continue;
        }
    }

    UpdateQueuePausedTasks();
    NotifyQueueChanged();
}

void FolderWindow::FileOperationState::CancelAll() noexcept
{
    std::vector<Task*> tasks;
    {
        std::scoped_lock lock(_mutex);
        tasks.reserve(_tasks.size());
        for (const auto& task : _tasks)
        {
            if (task)
            {
                tasks.push_back(task.get());
            }
        }
    }

    for (Task* task : tasks)
    {
        if (task)
        {
            task->RequestCancel();
        }
    }
}

void FolderWindow::FileOperationState::CollectTasks(std::vector<Task*>& outTasks) noexcept
{
    std::scoped_lock lock(_mutex);
    outTasks.clear();
    outTasks.reserve(_tasks.size());
    for (const auto& task : _tasks)
    {
        if (task)
        {
            outTasks.push_back(task.get());
        }
    }
}

void FolderWindow::FileOperationState::CollectCompletedTasks(std::vector<CompletedTaskSummary>& outTasks) noexcept
{
    std::scoped_lock lock(_mutex);
    outTasks.clear();
    outTasks.reserve(_completedTasks.size());
    for (const auto& summary : _completedTasks)
    {
        outTasks.push_back(summary);
    }
}

void FolderWindow::FileOperationState::DismissCompletedTask(uint64_t taskId) noexcept
{
    wil::unique_hwnd popupToClose;

    {
        std::scoped_lock lock(_mutex);
        _completedTasks.erase(std::remove_if(_completedTasks.begin(),
                                             _completedTasks.end(),
                                             [&](const CompletedTaskSummary& summary) noexcept { return summary.taskId == taskId; }),
                              _completedTasks.end());

        if (_tasks.empty() && _completedTasks.empty())
        {
            popupToClose = std::move(_popup);
        }
    }
}

bool FolderWindow::FileOperationState::GetAutoDismissSuccess() const noexcept
{
    if (! _owner._settings)
    {
        return false;
    }

    return GetAutoDismissSuccessFromSettings(*_owner._settings);
}

void FolderWindow::FileOperationState::SetAutoDismissSuccess(bool enabled) noexcept
{
    if (! _owner._settings)
    {
        return;
    }

    const bool previous = GetAutoDismissSuccessFromSettings(*_owner._settings);
    SetAutoDismissSuccessInSettings(*_owner._settings, enabled);

    if (enabled && ! previous)
    {
        wil::unique_hwnd popupToClose;
        {
            std::scoped_lock lock(_mutex);
            _completedTasks.erase(std::remove_if(_completedTasks.begin(),
                                                 _completedTasks.end(),
                                                 [](const CompletedTaskSummary& summary) noexcept
                                                 { return SUCCEEDED(summary.resultHr) || IsCancellationStatus(summary.resultHr); }),
                                  _completedTasks.end());

            if (_tasks.empty() && _completedTasks.empty())
            {
                popupToClose = std::move(_popup);
            }
        }
    }
}

bool FolderWindow::FileOperationState::OpenDiagnosticsLogForTask(uint64_t taskId) noexcept
{
    FlushDiagnostics(true);

    std::filesystem::path logPath;
    {
        std::scoped_lock lock(_mutex);
        for (const auto& summary : _completedTasks)
        {
            if (summary.taskId != taskId)
            {
                continue;
            }

            logPath = summary.diagnosticsLogPath;
            break;
        }
    }

    std::error_code ec;
    if (! logPath.empty() && ! std::filesystem::exists(logPath, ec))
    {
        logPath.clear();
    }

    if (logPath.empty())
    {
        std::scoped_lock lock(_mutex);
        logPath = GetLatestDiagnosticsLogPathUnlocked();
    }

    ec.clear();
    if (logPath.empty() || ! std::filesystem::exists(logPath, ec))
    {
        return false;
    }

    HINSTANCE hinst = ShellExecuteW(_owner.GetHwnd(), L"open", logPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(hinst) > 32;
}

bool FolderWindow::FileOperationState::ExportTaskIssuesReport(uint64_t taskId, std::filesystem::path* reportPathOut, bool openAfterExport) noexcept
{
    if (reportPathOut)
    {
        reportPathOut->clear();
    }

    FlushDiagnostics(true);

    CompletedTaskSummary summary{};
    bool found = false;
    {
        std::scoped_lock lock(_mutex);
        for (const auto& candidate : _completedTasks)
        {
            if (candidate.taskId != taskId)
            {
                continue;
            }

            summary = candidate;
            found   = true;
            break;
        }
    }

    if (! found)
    {
        return false;
    }

    if (summary.issueDiagnostics.empty() && summary.warningCount == 0 && summary.errorCount == 0)
    {
        return false;
    }

    const std::filesystem::path logsDir = GetDiagnosticsLogDirectory();
    if (logsDir.empty())
    {
        return false;
    }

    SYSTEMTIME localNow{};
    GetLocalTime(&localNow);

    wchar_t fileName[128]{};
    constexpr size_t fileNameMax                            = (sizeof(fileName) / sizeof(fileName[0])) - 1;
    const auto r                                            = std::format_to_n(fileName,
                                    fileNameMax,
                                    L"{}Task{}-{:04}{:02}{:02}-{:02}{:02}{:02}{:03}{}",
                                    kDiagnosticsIssueReportPrefix,
                                    static_cast<unsigned long long>(taskId),
                                    static_cast<unsigned>(localNow.wYear),
                                    static_cast<unsigned>(localNow.wMonth),
                                    static_cast<unsigned>(localNow.wDay),
                                    static_cast<unsigned>(localNow.wHour),
                                    static_cast<unsigned>(localNow.wMinute),
                                    static_cast<unsigned>(localNow.wSecond),
                                    static_cast<unsigned>(localNow.wMilliseconds),
                                    kDiagnosticsIssueReportExtension);
    fileName[(r.size < fileNameMax) ? r.size : fileNameMax] = L'\0';
    if (r.size > fileNameMax)
    {
        return false;
    }

    const std::filesystem::path reportPath = logsDir / fileName;

    std::error_code ec;
    std::filesystem::create_directories(logsDir, ec);
    if (ec)
    {
        return false;
    }

    wil::unique_handle file(CreateFileW(
        reportPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return false;
    }

    constexpr wchar_t bom = 0xFEFF;
    DWORD written         = 0;
    if (! WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr))
    {
        return false;
    }

    const std::wstring header =
        std::format(L"Task {:#x} ({})\r\nResult: 0x{:08X}\r\nWarnings: {:L}  Errors: {:L}\r\nCompleted items: {:L}/{:L}\r\nCompleted bytes: {:L}/{:L}\r\nFrom: "
                    L"{}\r\nTo: {}\r\n\r\nTime\tSeverity\tHRESULT\tStatus text\tCategory\tMessage\tSource\tDestination\r\n",
                    static_cast<unsigned long long>(summary.taskId),
                    OperationToString(summary.operation),
                    static_cast<unsigned long>(summary.resultHr),
                    summary.warningCount,
                    summary.errorCount,
                    summary.completedItems,
                    summary.totalItems,
                    summary.completedBytes,
                    summary.totalBytes,
                    summary.sourcePath.empty() ? L"-" : summary.sourcePath,
                    summary.destinationPath.empty() ? L"-" : summary.destinationPath);
    const size_t headerBytes = header.size() * sizeof(wchar_t);
    if (headerBytes > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
    {
        return false;
    }
    if (! WriteFile(file.get(), header.data(), static_cast<DWORD>(headerBytes), &written, nullptr))
    {
        return false;
    }

    for (const TaskDiagnosticEntry& issue : summary.issueDiagnostics)
    {
        const std::wstring statusText = EscapeDiagnosticField(FormatDiagnosticStatusText(issue.status));
        const std::wstring line       = std::format(L"{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}\t{}\t0x{:08X}\t{}\t{}\t{}\t{}\t{}\r\n",
                                              static_cast<unsigned>(issue.localTime.wYear),
                                              static_cast<unsigned>(issue.localTime.wMonth),
                                              static_cast<unsigned>(issue.localTime.wDay),
                                              static_cast<unsigned>(issue.localTime.wHour),
                                              static_cast<unsigned>(issue.localTime.wMinute),
                                              static_cast<unsigned>(issue.localTime.wSecond),
                                              static_cast<unsigned>(issue.localTime.wMilliseconds),
                                              DiagnosticSeverityToString(issue.severity),
                                              static_cast<unsigned long>(issue.status),
                                              statusText,
                                              EscapeDiagnosticField(issue.category),
                                              EscapeDiagnosticField(issue.message),
                                              EscapeDiagnosticField(issue.sourcePath),
                                              EscapeDiagnosticField(issue.destinationPath));

        const size_t bytesToWrite = line.size() * sizeof(wchar_t);
        if (bytesToWrite > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
        {
            continue;
        }

        if (! WriteFile(file.get(), line.data(), static_cast<DWORD>(bytesToWrite), &written, nullptr))
        {
            return false;
        }
    }

    if (reportPathOut)
    {
        *reportPathOut = reportPath;
    }

    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    CleanupDiagnosticsFilesInDirectory(
        logsDir, kDiagnosticsIssueReportPrefix, kDiagnosticsIssueReportExtension, diagnosticsSettings.maxDiagnosticsIssueReportFiles);

    if (! openAfterExport)
    {
        return true;
    }

    HINSTANCE hinst = ShellExecuteW(_owner.GetHwnd(), L"open", reportPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(hinst) > 32;
}

void FolderWindow::FileOperationState::ToggleIssuesPane() noexcept
{
    HWND pane = nullptr;
    {
        std::scoped_lock lock(_mutex);
        pane = _issuesPane.get();
    }

    if (pane)
    {
        if (IsWindowVisible(pane))
        {
            SaveIssuesPanePlacement(pane);
            ShowWindow(pane, SW_HIDE);
        }
        else
        {
            ShowWindow(pane, SW_SHOW);
            SetForegroundWindow(pane);
            PostMessageW(pane, WM_THEMECHANGED, 0, 0);
        }
        return;
    }

    HWND ownerWindow = _owner.GetHwnd();
    if (ownerWindow)
    {
        HWND rootWindow = GetAncestor(ownerWindow, GA_ROOT);
        if (rootWindow)
        {
            ownerWindow = rootWindow;
        }
    }

    if (! ownerWindow)
    {
        ownerWindow = GetParent(_owner.GetHwnd());
        if (! ownerWindow)
        {
            ownerWindow = _owner.GetHwnd();
        }
    }

    HWND createdPane = FileOperationsIssuesPane::Create(this, &_owner, ownerWindow);
    if (! createdPane)
    {
        return;
    }

    {
        std::scoped_lock lock(_mutex);
        _issuesPane.reset(createdPane);
    }
}

bool FolderWindow::FileOperationState::IsIssuesPaneVisible() noexcept
{
    std::scoped_lock lock(_mutex);
    return _issuesPane && IsWindowVisible(_issuesPane.get()) != FALSE;
}

bool FolderWindow::FileOperationState::TryGetIssuesPanePlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept
{
    outRect      = RECT{};
    outMaximized = false;

    if (! _owner._settings)
    {
        return false;
    }

    const std::wstring windowId(kFileOpsIssuesPaneWindowId);
    const auto it = _owner._settings->windows.find(windowId);
    if (it == _owner._settings->windows.end())
    {
        return false;
    }

    const Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(it->second, currentDpi);
    outRect.left                                       = normalized.bounds.x;
    outRect.top                                        = normalized.bounds.y;
    outRect.right                                      = normalized.bounds.x + std::max(1, normalized.bounds.width);
    outRect.bottom                                     = normalized.bounds.y + std::max(1, normalized.bounds.height);
    outMaximized                                       = normalized.state == Common::Settings::WindowState::Maximized;
    return true;
}

void FolderWindow::FileOperationState::SaveIssuesPanePlacement(HWND hwnd) noexcept
{
    if (! hwnd || ! _owner._settings || IsIconic(hwnd))
    {
        return;
    }

    WINDOWPLACEMENT placement{};
    placement.length = sizeof(placement);
    if (! GetWindowPlacement(hwnd, &placement))
    {
        return;
    }

    Common::Settings::WindowPlacement saved{};
    saved.state         = placement.showCmd == SW_SHOWMAXIMIZED ? Common::Settings::WindowState::Maximized : Common::Settings::WindowState::Normal;
    saved.bounds.x      = placement.rcNormalPosition.left;
    saved.bounds.y      = placement.rcNormalPosition.top;
    saved.bounds.width  = std::max(1, static_cast<int>(placement.rcNormalPosition.right - placement.rcNormalPosition.left));
    saved.bounds.height = std::max(1, static_cast<int>(placement.rcNormalPosition.bottom - placement.rcNormalPosition.top));
    saved.dpi           = GetDpiForWindow(hwnd);

    _owner._settings->windows[std::wstring(kFileOpsIssuesPaneWindowId)] = std::move(saved);
}

bool FolderWindow::FileOperationState::TryGetPopupPlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept
{
    outRect      = RECT{};
    outMaximized = false;

    if (! _owner._settings)
    {
        return false;
    }

    const std::wstring windowId(kFileOpsPopupWindowId);
    const auto it = _owner._settings->windows.find(windowId);
    if (it == _owner._settings->windows.end())
    {
        return false;
    }

    const Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(it->second, currentDpi);
    outRect.left                                       = normalized.bounds.x;
    outRect.top                                        = normalized.bounds.y;
    outRect.right                                      = normalized.bounds.x + std::max(1, normalized.bounds.width);
    outRect.bottom                                     = normalized.bounds.y + std::max(1, normalized.bounds.height);
    outMaximized                                       = normalized.state == Common::Settings::WindowState::Maximized;
    return true;
}

void FolderWindow::FileOperationState::SavePopupPlacement(HWND hwnd) noexcept
{
    if (! hwnd || ! _owner._settings || IsIconic(hwnd))
    {
        return;
    }

    WINDOWPLACEMENT placement{};
    placement.length = sizeof(placement);
    if (! GetWindowPlacement(hwnd, &placement))
    {
        return;
    }

    Common::Settings::WindowPlacement saved{};
    saved.state         = placement.showCmd == SW_SHOWMAXIMIZED ? Common::Settings::WindowState::Maximized : Common::Settings::WindowState::Normal;
    saved.bounds.x      = placement.rcNormalPosition.left;
    saved.bounds.y      = placement.rcNormalPosition.top;
    saved.bounds.width  = std::max(1, static_cast<int>(placement.rcNormalPosition.right - placement.rcNormalPosition.left));
    saved.bounds.height = std::max(1, static_cast<int>(placement.rcNormalPosition.bottom - placement.rcNormalPosition.top));
    saved.dpi           = GetDpiForWindow(hwnd);

    _owner._settings->windows[std::wstring(kFileOpsPopupWindowId)] = std::move(saved);
}

void FolderWindow::FileOperationState::OnPopupDestroyed(HWND hwnd) noexcept
{
    if (hwnd && _owner._settings)
    {
        SavePopupPlacement(hwnd);

        const Common::Settings::Settings settingsToSave = SettingsSave::PrepareForSave(*_owner._settings);
        const HRESULT saveHr                            = Common::Settings::SaveSettings(kFileOpsAppId, settingsToSave);
        if (FAILED(saveHr))
        {
            const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kFileOpsAppId);
            Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
        }
    }

    std::scoped_lock lock(_mutex);
    if (_popup.get() == hwnd)
    {
        static_cast<void>(_popup.release());
    }
}

void FolderWindow::FileOperationState::OnIssuesPaneDestroyed(HWND hwnd) noexcept
{
    if (hwnd && _owner._settings)
    {
        SaveIssuesPanePlacement(hwnd);

        const Common::Settings::Settings settingsToSave = SettingsSave::PrepareForSave(*_owner._settings);
        const HRESULT saveHr                            = Common::Settings::SaveSettings(kFileOpsAppId, settingsToSave);
        if (FAILED(saveHr))
        {
            const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kFileOpsAppId);
            Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
        }
    }

    std::scoped_lock lock(_mutex);
    if (_issuesPane.get() == hwnd)
    {
        static_cast<void>(_issuesPane.release());
    }
}

void FolderWindow::FileOperationState::UpdateLastPopupRect(const RECT& rect) noexcept
{
    std::scoped_lock lock(_mutex);
    _lastPopupRect = rect;
}

std::optional<RECT> FolderWindow::FileOperationState::GetLastPopupRect() noexcept
{
    std::scoped_lock lock(_mutex);
    return _lastPopupRect;
}

std::filesystem::path FolderWindow::FileOperationState::GetDiagnosticsLogDirectory() noexcept
{
    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kFileOpsAppId);
    if (settingsPath.empty())
    {
        return {};
    }

    const std::filesystem::path settingsDir = settingsPath.parent_path();
    if (settingsDir.empty())
    {
        return {};
    }

    const std::filesystem::path appRootDir = settingsDir.parent_path();
    if (appRootDir.empty())
    {
        return settingsDir / L"Logs";
    }

    // Keep diagnostics logs as a top-level sibling of Settings/Crashes.
    return appRootDir / L"Logs";
}

std::filesystem::path FolderWindow::FileOperationState::GetDiagnosticsLogPathForDate(const SYSTEMTIME& localTime) noexcept
{
    const std::filesystem::path logsDir = GetDiagnosticsLogDirectory();
    if (logsDir.empty())
    {
        return {};
    }

    wchar_t fileName[64]{};
    constexpr size_t fileNameMax                            = (sizeof(fileName) / sizeof(fileName[0])) - 1;
    const auto r                                            = std::format_to_n(fileName,
                                    fileNameMax,
                                    L"{}{:04}{:02}{:02}{}",
                                    kDiagnosticsLogPrefix,
                                    static_cast<unsigned>(localTime.wYear),
                                    static_cast<unsigned>(localTime.wMonth),
                                    static_cast<unsigned>(localTime.wDay),
                                    kDiagnosticsLogExtension);
    fileName[(r.size < fileNameMax) ? r.size : fileNameMax] = L'\0';
    if (r.size > fileNameMax)
    {
        return {};
    }

    return logsDir / fileName;
}

std::filesystem::path FolderWindow::FileOperationState::GetLatestDiagnosticsLogPathUnlocked() const noexcept
{
    const std::filesystem::path logsDir = GetDiagnosticsLogDirectory();
    if (logsDir.empty())
    {
        return {};
    }

    std::error_code ec;
    std::filesystem::path newestPath;
    for (std::filesystem::directory_iterator it(logsDir, ec), end; ! ec && it != end; it.increment(ec))
    {
        const auto& de = *it;
        if (! de.is_regular_file(ec))
        {
            continue;
        }

        const std::wstring fileNameText = de.path().filename().wstring();
        if (fileNameText.size() < (kDiagnosticsLogPrefix.size() + kDiagnosticsLogExtension.size()))
        {
            continue;
        }
        if (fileNameText.rfind(kDiagnosticsLogPrefix.data(), 0) != 0)
        {
            continue;
        }
        if (de.path().extension().wstring() != kDiagnosticsLogExtension)
        {
            continue;
        }

        if (newestPath.empty() || de.path().filename().wstring() > newestPath.filename().wstring())
        {
            newestPath = de.path();
        }
    }

    return newestPath;
}

void FolderWindow::FileOperationState::FlushDiagnostics(bool force) noexcept
{
    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    std::vector<TaskDiagnosticEntry> pending;
    ULONGLONG nowTick = GetTickCount64();

    {
        std::scoped_lock lock(_diagnosticsMutex);

        const bool flushIntervalReached =
            _lastDiagnosticsFlushTick == 0 ||
            (nowTick >= _lastDiagnosticsFlushTick && (nowTick - _lastDiagnosticsFlushTick) >= diagnosticsSettings.diagnosticsFlushIntervalMs);

        if (! force && ! flushIntervalReached && _diagnosticsPendingFlush.size() < diagnosticsSettings.maxDiagnosticsPerFlush)
        {
            return;
        }

        if (_diagnosticsPendingFlush.empty())
        {
            return;
        }

        pending.assign(std::make_move_iterator(_diagnosticsPendingFlush.begin()), std::make_move_iterator(_diagnosticsPendingFlush.end()));
        _diagnosticsPendingFlush.clear();
        _lastDiagnosticsFlushTick = nowTick;
    }

    const auto requeuePending = [&](size_t startIndex) noexcept
    {
        if (startIndex >= pending.size())
        {
            return;
        }

        std::scoped_lock lock(_diagnosticsMutex);
        const auto startIt = std::next(pending.begin(), static_cast<std::ptrdiff_t>(startIndex));
        _diagnosticsPendingFlush.insert(_diagnosticsPendingFlush.begin(), std::make_move_iterator(startIt), std::make_move_iterator(pending.end()));
    };

    SYSTEMTIME localNow{};
    GetLocalTime(&localNow);
    const std::filesystem::path logPath = GetDiagnosticsLogPathForDate(localNow);
    if (logPath.empty())
    {
        requeuePending(0);
        return;
    }

    const std::filesystem::path logsDir = logPath.parent_path();

    std::error_code ec;
    std::filesystem::create_directories(logsDir, ec);
    if (ec)
    {
        requeuePending(0);
        return;
    }

    wil::unique_handle file(CreateFileW(
        logPath.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        requeuePending(0);
        return;
    }

    LARGE_INTEGER fileSize{};
    bool shouldWriteBom = false;
    if (GetFileSizeEx(file.get(), &fileSize) && fileSize.QuadPart == 0)
    {
        shouldWriteBom = true;
    }

    if (SetFilePointer(file.get(), 0, nullptr, FILE_END) == INVALID_SET_FILE_POINTER && GetLastError() != NO_ERROR)
    {
        requeuePending(0);
        return;
    }

    if (shouldWriteBom)
    {
        constexpr wchar_t bom = 0xFEFF;
        DWORD written         = 0;
        if (! WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr) || written != sizeof(bom))
        {
            requeuePending(0);
            return;
        }
    }

    for (size_t index = 0; index < pending.size(); ++index)
    {
        const TaskDiagnosticEntry& entry = pending[index];
        const wchar_t* categoryText      = entry.category.empty() ? L"general" : entry.category.c_str();
        const unsigned long hrU32        = static_cast<unsigned long>(entry.status);

        const std::wstring timeText = std::format(L"{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}",
                                                  static_cast<unsigned>(entry.localTime.wYear),
                                                  static_cast<unsigned>(entry.localTime.wMonth),
                                                  static_cast<unsigned>(entry.localTime.wDay),
                                                  static_cast<unsigned>(entry.localTime.wHour),
                                                  static_cast<unsigned>(entry.localTime.wMinute),
                                                  static_cast<unsigned>(entry.localTime.wSecond),
                                                  static_cast<unsigned>(entry.localTime.wMilliseconds));
        const std::wstring hrHex    = std::format(L"0x{:08X}", hrU32);

        const std::wstring escapedCategory = EscapeDiagnosticJsonString(categoryText);
        const std::wstring escapedMessage  = EscapeDiagnosticJsonString(entry.message);
        const std::wstring escapedSource   = EscapeDiagnosticJsonString(entry.sourcePath);
        const std::wstring escapedDest     = EscapeDiagnosticJsonString(entry.destinationPath);
        const std::wstring escapedHrName   = EscapeDiagnosticJsonString(FormatDiagnosticHresultName(entry.status));

        std::wstring escapedHrText;
        if (entry.status != S_OK)
        {
            escapedHrText = EscapeDiagnosticJsonString(FormatDiagnosticStatusText(entry.status));
        }

        std::wstring line;
        line.reserve(256u + timeText.size() + hrHex.size() + escapedHrName.size() + escapedCategory.size() + escapedMessage.size() + escapedSource.size() +
                     escapedDest.size() + escapedHrText.size());

        line.append(L"{\"ts\":\"");
        line.append(timeText);
        line.append(L"\",\"level\":\"");
        line.append(DiagnosticSeverityToString(entry.severity));
        line.append(L"\",\"task\":");
        line.append(std::to_wstring(static_cast<unsigned long long>(entry.taskId)));
        line.append(L",\"op\":\"");
        line.append(OperationToString(entry.operation));
        line.append(L"\",\"category\":\"");
        line.append(escapedCategory);
        line.append(L"\",\"hr\":\"");
        line.append(hrHex);
        line.append(L"\",\"hrName\":\"");
        line.append(escapedHrName);
        line.append(L"\"");
        if (! escapedHrText.empty())
        {
            line.append(L",\"hrText\":\"");
            line.append(escapedHrText);
            line.append(L"\"");
        }
        if (entry.processWorkingSetBytes != 0 || entry.processPrivateBytes != 0)
        {
            line.append(L",\"memWorkingSetBytes\":");
            line.append(std::to_wstring(entry.processWorkingSetBytes));
            line.append(L",\"memPrivateBytes\":");
            line.append(std::to_wstring(entry.processPrivateBytes));
        }
        line.append(L",\"message\":\"");
        line.append(escapedMessage);
        line.append(L"\"");

        if (! entry.sourcePath.empty())
        {
            line.append(L",\"src\":\"");
            line.append(escapedSource);
            line.append(L"\"");

            const std::wstring_view leaf = GetPathLeaf(entry.sourcePath);
            if (! leaf.empty())
            {
                line.append(L",\"srcLeaf\":\"");
                line.append(EscapeDiagnosticJsonString(leaf));
                line.append(L"\"");
            }
        }
        else
        {
            line.append(L",\"src\":null");
        }

        if (! entry.destinationPath.empty())
        {
            line.append(L",\"dst\":\"");
            line.append(escapedDest);
            line.append(L"\"");

            const std::wstring_view leaf = GetPathLeaf(entry.destinationPath);
            if (! leaf.empty())
            {
                line.append(L",\"dstLeaf\":\"");
                line.append(EscapeDiagnosticJsonString(leaf));
                line.append(L"\"");
            }
        }
        else
        {
            line.append(L",\"dst\":null");
        }

        line.append(L"}\r\n");

        const size_t bytesToWrite = line.size() * sizeof(wchar_t);
        if (bytesToWrite > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
        {
            continue;
        }

        DWORD written = 0;
        if (! WriteFile(file.get(), line.data(), static_cast<DWORD>(bytesToWrite), &written, nullptr) || written != static_cast<DWORD>(bytesToWrite))
        {
            requeuePending(index);
            return;
        }
    }

    bool runCleanup = force;
    {
        std::scoped_lock lock(_diagnosticsMutex);
        if (! runCleanup)
        {
            runCleanup = _lastDiagnosticsCleanupTick == 0 || (nowTick >= _lastDiagnosticsCleanupTick &&
                                                              (nowTick - _lastDiagnosticsCleanupTick) >= diagnosticsSettings.diagnosticsCleanupIntervalMs);
        }
        if (runCleanup)
        {
            _lastDiagnosticsCleanupTick = nowTick;
        }
    }

    if (! runCleanup)
    {
        return;
    }

    CleanupDiagnosticsFilesInDirectory(logsDir, kDiagnosticsLogPrefix, kDiagnosticsLogExtension, diagnosticsSettings.maxDiagnosticsLogFiles);
    CleanupDiagnosticsFilesInDirectory(
        logsDir, kDiagnosticsIssueReportPrefix, kDiagnosticsIssueReportExtension, diagnosticsSettings.maxDiagnosticsIssueReportFiles);
}

void FolderWindow::FileOperationState::RecordTaskDiagnostic(uint64_t taskId,
                                                            FileSystemOperation operation,
                                                            DiagnosticSeverity severity,
                                                            HRESULT status,
                                                            std::wstring_view category,
                                                            std::wstring_view message,
                                                            std::wstring_view sourcePath,
                                                            std::wstring_view destinationPath) noexcept
{
    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    if (severity == DiagnosticSeverity::Info && ! diagnosticsSettings.infoEnabled)
    {
        return;
    }
    if (severity == DiagnosticSeverity::Debug && ! diagnosticsSettings.debugEnabled)
    {
        return;
    }

    TaskDiagnosticEntry entry{};
    GetLocalTime(&entry.localTime);
    entry.taskId    = taskId;
    entry.operation = operation;
    entry.severity  = severity;
    entry.status    = status;
    if (severity == DiagnosticSeverity::Debug || severity == DiagnosticSeverity::Error)
    {
        const ProcessMemorySnapshot snapshot = CaptureProcessMemorySnapshot();
        entry.processWorkingSetBytes         = snapshot.workingSetBytes;
        entry.processPrivateBytes            = snapshot.privateBytes;
    }
    entry.category        = std::wstring(category);
    entry.message         = std::wstring(message);
    entry.sourcePath      = std::wstring(sourcePath);
    entry.destinationPath = std::wstring(destinationPath);

    const ULONGLONG nowTick = GetTickCount64();
    bool shouldFlush        = false;
    {
        std::scoped_lock lock(_diagnosticsMutex);

        _diagnosticsInMemory.push_back(entry);
        while (_diagnosticsInMemory.size() > diagnosticsSettings.maxDiagnosticsInMemory)
        {
            _diagnosticsInMemory.pop_front();
        }

        _diagnosticsPendingFlush.push_back(entry);

        if (severity == DiagnosticSeverity::Warning || severity == DiagnosticSeverity::Error)
        {
            auto& counts = _taskDiagnosticCounts[taskId];
            if (severity == DiagnosticSeverity::Warning)
            {
                ++counts.first;
            }
            else
            {
                ++counts.second;
            }

            if (! message.empty())
            {
                _taskLastDiagnosticMessage[taskId] = std::wstring(message);
            }

            auto& issues = _taskIssueDiagnostics[taskId];
            issues.push_back(entry);
            while (issues.size() > kMaxTaskIssueDiagnostics)
            {
                issues.pop_front();
            }
        }

        const bool flushIntervalReached =
            _lastDiagnosticsFlushTick == 0 ||
            (nowTick >= _lastDiagnosticsFlushTick && (nowTick - _lastDiagnosticsFlushTick) >= diagnosticsSettings.diagnosticsFlushIntervalMs);
        shouldFlush = flushIntervalReached || _diagnosticsPendingFlush.size() >= diagnosticsSettings.maxDiagnosticsPerFlush;
    }

    if (shouldFlush)
    {
        FlushDiagnostics(false);
    }
}

void FolderWindow::FileOperationState::RecordCompletedTask(Task& task) noexcept
{
    CompletedTaskSummary summary{};
    SYSTEMTIME localNow{};
    GetLocalTime(&localNow);
    summary.taskId             = task._taskId;
    summary.operation          = task._operation;
    summary.sourcePane         = task._sourcePane;
    summary.destinationPane    = task._destinationPane;
    summary.destinationFolder  = task.GetDestinationFolder();
    summary.diagnosticsLogPath = GetDiagnosticsLogPathForDate(localNow);
    summary.resultHr           = task.GetResult();
    summary.completedTick      = GetTickCount64();

    {
        std::scoped_lock lock(task._progressMutex);
        summary.totalItems      = task._progressTotalItems;
        summary.completedItems  = task._progressCompletedItems;
        summary.totalBytes      = task._progressTotalBytes;
        summary.completedBytes  = task._progressCompletedBytes;
        summary.sourcePath      = task._progressSourcePath;
        summary.destinationPath = task._progressDestinationPath;
    }

    {
        std::scoped_lock lock(_diagnosticsMutex);
        const auto countsIt = _taskDiagnosticCounts.find(summary.taskId);
        if (countsIt != _taskDiagnosticCounts.end())
        {
            summary.warningCount = countsIt->second.first;
            summary.errorCount   = countsIt->second.second;
            _taskDiagnosticCounts.erase(countsIt);
        }

        const auto messageIt = _taskLastDiagnosticMessage.find(summary.taskId);
        if (messageIt != _taskLastDiagnosticMessage.end())
        {
            summary.lastDiagnosticMessage = messageIt->second;
            _taskLastDiagnosticMessage.erase(messageIt);
        }

        const auto issuesIt = _taskIssueDiagnostics.find(summary.taskId);
        if (issuesIt != _taskIssueDiagnostics.end())
        {
            summary.issueDiagnostics.assign(issuesIt->second.begin(), issuesIt->second.end());
            _taskIssueDiagnostics.erase(issuesIt);
        }
    }

    if (FAILED(summary.resultHr) && summary.warningCount == 0 && summary.errorCount == 0)
    {
        const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (summary.resultHr == partialHr)
        {
            summary.warningCount = 1;
            if (summary.lastDiagnosticMessage.empty())
            {
                summary.lastDiagnosticMessage = L"Task completed with skipped items.";
            }
        }
        else if (! IsCancellationStatus(summary.resultHr))
        {
            summary.errorCount = 1;
            if (summary.lastDiagnosticMessage.empty())
            {
                summary.lastDiagnosticMessage =
                    std::format(L"Task failed (0x{:08X}) without detailed diagnostics.", static_cast<unsigned long>(summary.resultHr));
            }
        }
    }

    if ((summary.warningCount > 0 || summary.errorCount > 0) && summary.issueDiagnostics.empty())
    {
        TaskDiagnosticEntry synthetic{};
        synthetic.localTime       = localNow;
        synthetic.taskId          = summary.taskId;
        synthetic.operation       = summary.operation;
        synthetic.severity        = summary.errorCount > 0 ? DiagnosticSeverity::Error : DiagnosticSeverity::Warning;
        synthetic.status          = summary.resultHr;
        synthetic.category        = L"task.summary";
        synthetic.message         = summary.lastDiagnosticMessage.empty() ? L"Task completed with diagnostics." : summary.lastDiagnosticMessage;
        synthetic.sourcePath      = summary.sourcePath;
        synthetic.destinationPath = summary.destinationPath;
        summary.issueDiagnostics.push_back(std::move(synthetic));
    }

    std::wstring completedStatus = L"success";
    if (summary.resultHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        completedStatus = L"partial";
    }
    else if (IsCancellationStatus(summary.resultHr))
    {
        completedStatus = L"canceled";
    }
    else if (FAILED(summary.resultHr))
    {
        completedStatus = L"failed";
    }

    const std::wstring completedMessage = std::format(L"Task completed: status={}, op={}, hr=0x{:08X}, items={:L}/{:L}, bytes={:L}/{:L}.",
                                                      completedStatus,
                                                      OperationToString(summary.operation),
                                                      static_cast<unsigned long>(summary.resultHr),
                                                      summary.completedItems,
                                                      summary.totalItems,
                                                      summary.completedBytes,
                                                      summary.totalBytes);
    RecordTaskDiagnostic(summary.taskId,
                         summary.operation,
                         DiagnosticSeverity::Info,
                         summary.resultHr,
                         L"task.completed",
                         completedMessage,
                         summary.sourcePath,
                         summary.destinationPath);

    {
        std::scoped_lock lock(_mutex);
        _completedTasks.push_front(std::move(summary));
        while (_completedTasks.size() > kMaxCompletedTaskSummaries)
        {
            _completedTasks.pop_back();
        }
    }

    FlushDiagnostics(true);
}

#ifdef _DEBUG
HWND FolderWindow::FileOperationState::GetPopupHwndForSelfTest() noexcept
{
    std::scoped_lock lock(_mutex);
    return _popup.get();
}
#endif

bool FolderWindow::FileOperationState::EnterOperation(Task& task, std::stop_token stopToken) noexcept
{
    std::unique_lock lock(_queueMutex);

    const bool waitForOthers = task._waitForOthers.load(std::memory_order_acquire);
    if (! waitForOthers)
    {
        ++_activeOperations;
        return true;
    }

    _queue.push_back(task._taskId);
    _queueCv.notify_all();

    _queueCv.wait(lock,
                  [&]
                  {
                      if (stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire))
                      {
                          return true;
                      }

                      if (! task._waitForOthers.load(std::memory_order_acquire))
                      {
                          return true;
                      }

                      return _activeOperations == 0 && ! _queue.empty() && _queue.front() == task._taskId;
                  });

    if (stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire))
    {
        RemoveFromQueue(task._taskId);
        return false;
    }

    if (! task._waitForOthers.load(std::memory_order_acquire))
    {
        RemoveFromQueue(task._taskId);
        ++_activeOperations;
        return true;
    }

    if (! _queue.empty() && _queue.front() == task._taskId)
    {
        _queue.pop_front();
    }
    ++_activeOperations;
    return true;
}

void FolderWindow::FileOperationState::LeaveOperation() noexcept
{
    {
        std::scoped_lock lock(_queueMutex);
        if (_activeOperations > 0)
        {
            --_activeOperations;
        }
    }
    _queueCv.notify_all();
}

void FolderWindow::FileOperationState::PostCompleted(Task& task) noexcept
{
    RecordCompletedTask(task);

    HWND owner = _owner.GetHwnd();
    if (! owner)
    {
        return;
    }

    auto payload    = std::make_unique<TaskCompletedPayload>();
    payload->taskId = task._taskId;
    payload->hr     = task.GetResult();

    static_cast<void>(PostMessagePayload(owner, WndMsg::kFileOperationCompleted, 0, std::move(payload)));
}

FolderWindow::FileOperationState::Task* FolderWindow::FileOperationState::FindTask(uint64_t taskId) noexcept
{
    std::scoped_lock lock(_mutex);
    for (auto& task : _tasks)
    {
        if (task && task->GetId() == taskId)
        {
            return task.get();
        }
    }
    return nullptr;
}

void FolderWindow::FileOperationState::RemoveTask(uint64_t taskId) noexcept
{
    wil::unique_hwnd popupToClose;
    bool shouldUpdateQueue = false;
    {
        std::scoped_lock lock(_mutex);

        _tasks.erase(std::remove_if(_tasks.begin(), _tasks.end(), [&](const std::unique_ptr<Task>& t) { return ! t || t->GetId() == taskId; }), _tasks.end());

        if (_tasks.empty() && _completedTasks.empty())
        {
            popupToClose = std::move(_popup);
        }
        else
        {
            shouldUpdateQueue = _queueNewTasks.load(std::memory_order_acquire);
        }
    }

    if (shouldUpdateQueue)
    {
        UpdateQueuePausedTasks();
    }
}

void FolderWindow::FileOperationState::RemoveFromQueue(uint64_t taskId) noexcept
{
    auto it = std::find(_queue.begin(), _queue.end(), taskId);
    if (it != _queue.end())
    {
        _queue.erase(it);
    }
}

void FolderWindow::FileOperationState::UpdateQueuePausedTasks() noexcept
{
    const bool queueMode = _queueNewTasks.load(std::memory_order_acquire);

    std::vector<Task*> tasks;
    CollectTasks(tasks);

    if (! queueMode)
    {
        for (auto* task : tasks)
        {
            if (task)
            {
                task->SetQueuePaused(false);
            }
        }
        return;
    }

    std::optional<uint64_t> firstActiveId;
    ULONGLONG firstActiveTick = std::numeric_limits<ULONGLONG>::max();
    for (auto* task : tasks)
    {
        if (! task)
        {
            continue;
        }

        if (! task->HasEnteredOperation())
        {
            continue;
        }

        const ULONGLONG enteredTick = task->GetEnteredOperationTick();
        const ULONGLONG tickKey     = enteredTick != 0 ? enteredTick : std::numeric_limits<ULONGLONG>::max();

        const uint64_t id = task->GetId();
        if (! firstActiveId.has_value() || tickKey < firstActiveTick || (tickKey == firstActiveTick && id < firstActiveId.value()))
        {
            firstActiveId   = id;
            firstActiveTick = tickKey;
        }
    }

    for (auto* task : tasks)
    {
        if (! task)
        {
            continue;
        }

        if (! task->HasEnteredOperation())
        {
            task->SetQueuePaused(false);
            continue;
        }

        const uint64_t id        = task->GetId();
        const bool isFirstActive = firstActiveId.has_value() && id == firstActiveId.value();
        task->SetQueuePaused(! isFirstActive);
    }
}
