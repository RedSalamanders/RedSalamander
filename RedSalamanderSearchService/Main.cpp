#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <sddl.h>
#include <shellapi.h>

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cstdint>
#include <cwchar>
#include <deque>
#include <format>
#include <mutex>
#include <string>
#include <string_view>
#include <thread>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"
#include "SearchServiceBroker.h"
#include "SqliteIndexStore.h"

#include <ftxui/component/component.hpp>
#include <ftxui/component/component_options.hpp>
#include <ftxui/component/event.hpp>
#include <ftxui/component/screen_interactive.hpp>
#include <ftxui/dom/elements.hpp>

#pragma comment(lib, "Advapi32.lib")

namespace
{
enum class StartupAction : uint8_t
{
    ServiceDispatcher = 0u,
    RunForeground,
    CompactStore,
    RequestCompactService,
    ShowHelp,
    RegisterService,
    UnregisterService,
};

struct ParsedArguments final
{
    StartupAction action                                          = StartupAction::ServiceDispatcher;
    LocalSearchIndexCore::PersistentStoreKind persistentStoreKind = LocalSearchIndexCore::PersistentStoreKind::Sqlite;
    uint32_t protocolVersion                                      = SearchServiceBroker::kProtocolVersion;
    uint32_t maxRequestsBeforeExit                                = 0u;
    uint32_t disconnectAfterBatches                               = 0u;
    std::wstring pipeName;
    std::wstring storageRootDirectory;
    std::wstring sqliteDatabasePath;
    std::wstring actionOption;
    std::wstring errorMessage;
    bool storeBackendExplicit = false;
};

struct CommandResult final
{
    HRESULT hr = S_OK;
    std::wstring message;
};

struct ForegroundRequestSnapshot final
{
    SearchServiceBroker::ServerRequestType requestType          = SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE;
    FileSystemSearchPhase requestPhase                          = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    uint32_t warningFlags                                       = FILESYSTEM_SEARCH_WARNING_NONE;
    uint32_t requestFlags                                       = FILESYSTEM_SEARCH_NONE;
    uint32_t requestNameMode                                    = FILESYSTEM_SEARCH_NAME_DISABLED;
    LocalSearchIndexCore::StoreState storeState                 = LocalSearchIndexCore::StoreState::Unknown;
    LocalSearchIndexCore::SyncPhase syncPhase                   = LocalSearchIndexCore::SyncPhase::Idle;
    LocalSearchIndexCore::QueryExecutionMode queryExecutionMode = LocalSearchIndexCore::QueryExecutionMode::Unknown;
    LocalSearchIndexCore::FallbackReason fallbackReason         = LocalSearchIndexCore::FallbackReason::None;
    uint64_t completedRoots                                     = 0u;
    uint64_t totalRoots                                         = 0u;
    uint32_t batchesSent                                        = 0u;
    uint64_t scannedDirectories                                 = 0u;
    uint64_t scannedFiles                                       = 0u;
    uint64_t candidateFiles                                     = 0u;
    uint64_t matchedEntries                                     = 0u;
    uint64_t snapshotFileBytes                                  = 0u;
    uint64_t estimatedMemoryBytes                               = 0u;
    uint64_t ensureReadyDurationMs                              = 0u;
    uint64_t executeQueryDurationMs                             = 0u;
    std::wstring rootPath;
    std::wstring namePattern;
    std::wstring activeRoot;
    std::wstring currentPath;
};

struct ForegroundWarmupSnapshot final
{
    bool enabled            = false;
    bool running            = false;
    uint32_t totalRoots     = 0u;
    uint32_t completedRoots = 0u;
    uint32_t failedRoots    = 0u;
    bool hasFailure         = false;
    HRESULT lastFailureHr   = S_OK;
    std::wstring currentRoot;
    std::wstring lastFailureRoot;
};

struct ForegroundEventRecord final
{
    SearchServiceBroker::ServerEventType type = SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING;
    HRESULT hr                                = S_OK;
    uint64_t tickMs                           = 0u;
    uint32_t totalConnections                 = 0u;
    uint32_t handledRequests                  = 0u;
    ForegroundRequestSnapshot requestSnapshot{};
    ForegroundWarmupSnapshot warmupSnapshot{};
};

struct ForegroundConsoleState final
{
    ForegroundConsoleState()                                         = default;
    ForegroundConsoleState(const ForegroundConsoleState&)            = delete;
    ForegroundConsoleState& operator=(const ForegroundConsoleState&) = delete;
    ForegroundConsoleState(ForegroundConsoleState&&)                 = delete;
    ForegroundConsoleState& operator=(ForegroundConsoleState&&)      = delete;

    std::atomic<uint32_t> lastEvent{SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING};
    std::atomic<long> lastHr{S_OK};
    std::atomic<uint32_t> handledRequests{0u};
    std::atomic<uint32_t> totalConnections{0u};
    std::atomic<uint32_t> lastActivityEvent{SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING};
    std::atomic<long> lastActivityHr{S_OK};
    std::atomic<uint32_t> lastActivityRequests{0u};
    std::atomic<uint64_t> lastActivityTickMs{0u};
    std::atomic<uint64_t> activeRequestStartedTickMs{0u};
    std::atomic<uint64_t> uiRevision{0u};
    std::atomic<uint32_t> seenEventMask{0u};
    std::atomic<bool> maintenanceQueued{false};
    std::atomic<bool> maintenanceRunning{false};
    std::atomic<bool> stopRequested{false};
    mutable std::mutex detailsMutex;
    bool startupWarmupEnabled                          = false;
    bool startupWarmupRunning                          = false;
    uint32_t startupWarmupTotalRoots                   = 0u;
    uint32_t startupWarmupCompletedRoots               = 0u;
    uint32_t startupWarmupFailedRoots                  = 0u;
    bool startupWarmupHasFailure                       = false;
    HRESULT startupWarmupLastFailureHr                 = S_OK;
    SearchServiceBroker::ServerRequestType requestType = SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE;
    FileSystemSearchPhase requestPhase                 = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    uint32_t requestFlags                              = FILESYSTEM_SEARCH_NONE;
    uint32_t requestNameMode                           = FILESYSTEM_SEARCH_NAME_DISABLED;
    uint32_t warningFlags                              = FILESYSTEM_SEARCH_WARNING_NONE;
    uint32_t storeState                                = static_cast<uint32_t>(LocalSearchIndexCore::StoreState::Unknown);
    uint32_t syncPhase                                 = static_cast<uint32_t>(LocalSearchIndexCore::SyncPhase::Idle);
    uint32_t queryExecutionMode                        = static_cast<uint32_t>(LocalSearchIndexCore::QueryExecutionMode::Unknown);
    uint32_t fallbackReason                            = static_cast<uint32_t>(LocalSearchIndexCore::FallbackReason::None);
    uint64_t completedRoots                            = 0u;
    uint64_t totalRoots                                = 0u;
    uint32_t batchesSent                               = 0u;
    uint64_t scannedDirectories                        = 0u;
    uint64_t scannedFiles                              = 0u;
    uint64_t candidateFiles                            = 0u;
    uint64_t matchedEntries                            = 0u;
    uint64_t snapshotFileBytes                         = 0u;
    uint64_t estimatedMemoryBytes                      = 0u;
    uint64_t ensureReadyDurationMs                     = 0u;
    uint64_t executeQueryDurationMs                    = 0u;
    std::array<wchar_t, 320> rootPath{};
    std::array<wchar_t, 160> namePattern{};
    std::array<wchar_t, 320> activeRoot{};
    std::array<wchar_t, 320> currentPath{};
    std::array<wchar_t, 320> startupWarmupCurrentRoot{};
    std::array<wchar_t, 320> startupWarmupLastFailureRoot{};
    std::deque<ForegroundEventRecord> pendingEvents;
};

struct ForegroundConsoleSession final
{
    ForegroundConsoleSession()                                           = default;
    ForegroundConsoleSession(const ForegroundConsoleSession&)            = delete;
    ForegroundConsoleSession& operator=(const ForegroundConsoleSession&) = delete;
    ForegroundConsoleSession(ForegroundConsoleSession&&)                 = delete;
    ForegroundConsoleSession& operator=(ForegroundConsoleSession&&)      = delete;

    wil::unique_handle stdinHandle;
    wil::unique_handle stdoutHandle;
    wil::unique_handle stderrHandle;
    DWORD originalOutputMode = 0u;
    bool renderDashboard     = false;
    bool restoreOutputMode   = false;
    bool useAnsiColors       = false;
    bool useAlternateScreen  = false;
    bool cursorHidden        = false;

    [[nodiscard]] HANDLE GetOutputHandle() const noexcept
    {
        return stdoutHandle ? stdoutHandle.get() : ::GetStdHandle(STD_OUTPUT_HANDLE);
    }
};

struct ForegroundDashboardSnapshot final
{
    SearchServiceBroker::ServerEventType eventType         = SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING;
    SearchServiceBroker::ServerEventType lastActivityEvent = SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING;
    HRESULT lastHr                                         = S_OK;
    HRESULT lastActivityHr                                 = S_OK;
    uint32_t handledRequests                               = 0u;
    uint32_t totalConnections                              = 0u;
    uint64_t lastActivityTickMs                            = 0u;
    uint64_t activeRequestStartedTickMs                    = 0u;
    uint64_t uiRevision                                    = 0u;
    bool stopRequested                                     = false;
    ForegroundRequestSnapshot requestSnapshot{};
    std::vector<ForegroundEventRecord> eventHistory;
};

SERVICE_STATUS_HANDLE g_serviceStatusHandle = nullptr;
SERVICE_STATUS g_serviceStatus{};
wil::unique_event g_serviceStopEvent;
std::atomic<HANDLE> g_foregroundStopEvent{nullptr};
std::atomic<HANDLE> g_foregroundUiWakeEvent{nullptr};
std::atomic<ForegroundConsoleState*> g_foregroundConsoleState{nullptr};
constexpr wchar_t kSingleInstanceEventNameEnvVar[] = L"REDSALAMANDER_SEARCH_SERVICE_INSTANCE_EVENT";

void UpdateServiceStatus(DWORD currentState, DWORD win32ExitCode, DWORD waitHint) noexcept
{
    if (! g_serviceStatusHandle)
    {
        return;
    }

    g_serviceStatus.dwServiceType      = SERVICE_WIN32_OWN_PROCESS;
    g_serviceStatus.dwCurrentState     = currentState;
    g_serviceStatus.dwWin32ExitCode    = win32ExitCode;
    g_serviceStatus.dwWaitHint         = waitHint;
    g_serviceStatus.dwControlsAccepted = currentState == SERVICE_START_PENDING ? 0u : SERVICE_ACCEPT_STOP | SERVICE_ACCEPT_SHUTDOWN;
    g_serviceStatus.dwCheckPoint       = (currentState == SERVICE_RUNNING || currentState == SERVICE_STOPPED) ? 0u : 1u;
    static_cast<void>(::SetServiceStatus(g_serviceStatusHandle, &g_serviceStatus));
}

[[nodiscard]] std::wstring GetEnvironmentVariableText(std::wstring_view name) noexcept
{
    if (name.empty())
    {
        return {};
    }

    std::wstring key(name);
    const DWORD required = ::GetEnvironmentVariableW(key.c_str(), nullptr, 0u);
    if (required == 0u)
    {
        return {};
    }

    std::wstring value(required, L'\0');
    const DWORD written = ::GetEnvironmentVariableW(key.c_str(), value.data(), required);
    if (written == 0u || written >= required)
    {
        return {};
    }

    value.resize(written);
    return value;
}

[[nodiscard]] std::wstring GetSingleInstanceEventName() noexcept
{
    std::wstring eventName = GetEnvironmentVariableText(kSingleInstanceEventNameEnvVar);
    if (! eventName.empty())
    {
        return eventName;
    }

    eventName.assign(L"Global\\");
    eventName.append(SearchServiceBroker::kServiceName);
    eventName.append(L".Instance");
    return eventName;
}

HRESULT CreateSingleInstanceSecurity(SECURITY_ATTRIBUTES& outAttributes, wil::unique_hlocal& outDescriptor) noexcept
{
    outAttributes = {};
    outDescriptor.reset();

    static constexpr wchar_t kSingleInstanceSddl[] = L"D:P(A;;GA;;;SY)(A;;GA;;;BA)(A;;GA;;;IU)";

    PSECURITY_DESCRIPTOR descriptor = nullptr;
    if (::ConvertStringSecurityDescriptorToSecurityDescriptorW(kSingleInstanceSddl, SDDL_REVISION_1, &descriptor, nullptr) == 0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    outDescriptor.reset(descriptor);
    outAttributes.nLength              = sizeof(outAttributes);
    outAttributes.lpSecurityDescriptor = outDescriptor.get();
    outAttributes.bInheritHandle       = FALSE;
    return S_OK;
}

HRESULT AcquireSingleInstanceGuard(wil::unique_handle& outGuard, bool& outAlreadyRunning) noexcept
{
    outGuard.reset();
    outAlreadyRunning = false;

    SECURITY_ATTRIBUTES attributes{};
    wil::unique_hlocal descriptor;
    HRESULT hr = CreateSingleInstanceSecurity(attributes, descriptor);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring eventName = GetSingleInstanceEventName();
    ::SetLastError(ERROR_SUCCESS);
    outGuard.reset(::CreateEventW(&attributes, TRUE, FALSE, eventName.c_str()));
    if (! outGuard)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    outAlreadyRunning = (::GetLastError() == ERROR_ALREADY_EXISTS);
    return S_OK;
}

[[nodiscard]] std::wstring BuildAlreadyRunningMessage() noexcept
{
    return std::format(L"Another '{}' instance is already running. Stop the existing Windows service or foreground process before starting a new one.",
                       SearchServiceBroker::kServiceName);
}

[[nodiscard]] LocalSearchIndexCore::RepositoryOptions BuildRepositoryOptions(const ParsedArguments& parsed) noexcept
{
    return {
        .snapshotRootDirectory = parsed.storageRootDirectory,
        .persistentStoreKind   = parsed.persistentStoreKind,
        .sqliteDatabasePath    = parsed.sqliteDatabasePath,
        .sqliteAuthoritative   = parsed.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
    };
}

[[nodiscard]] bool TryParseStoreBackend(std::wstring_view text, LocalSearchIndexCore::PersistentStoreKind& outKind) noexcept
{
    if (OrdinalString::EqualsNoCase(text, L"snapshot") || OrdinalString::EqualsNoCase(text, L"snapshot-v1"))
    {
        outKind = LocalSearchIndexCore::PersistentStoreKind::SnapshotBinary;
        return true;
    }

    if (OrdinalString::EqualsNoCase(text, L"sqlite") || OrdinalString::EqualsNoCase(text, L"sqlite-v2"))
    {
        outKind = LocalSearchIndexCore::PersistentStoreKind::Sqlite;
        return true;
    }

    return false;
}

[[nodiscard]] bool TryParseUnsigned(std::wstring_view text, uint32_t& outValue) noexcept
{
    outValue = 0u;
    if (text.empty())
    {
        return false;
    }

    uint64_t value = 0u;
    for (const wchar_t ch : text)
    {
        if (ch < L'0' || ch > L'9')
        {
            return false;
        }

        value = (value * 10u) + static_cast<uint64_t>(ch - L'0');
        if (value > UINT32_MAX)
        {
            return false;
        }
    }

    outValue = static_cast<uint32_t>(value);
    return true;
}

[[nodiscard]] std::wstring_view GetActionOptionName(StartupAction action) noexcept
{
    switch (action)
    {
        case StartupAction::RunForeground: return L"--run-foreground";
        case StartupAction::CompactStore: return L"--compact";
        case StartupAction::RequestCompactService: return L"--request-compact";
        case StartupAction::ShowHelp: return L"--help";
        case StartupAction::RegisterService: return L"--register";
        case StartupAction::UnregisterService: return L"--unregister";
        case StartupAction::ServiceDispatcher:
        default: return L"<service>";
    }
}

[[nodiscard]] bool TrySetAction(ParsedArguments& parsed, StartupAction action, std::wstring_view option) noexcept
{
    if (action == StartupAction::ShowHelp)
    {
        parsed.action       = action;
        parsed.actionOption = std::wstring(option);
        return true;
    }

    if (parsed.action == StartupAction::ShowHelp)
    {
        return true;
    }

    if (parsed.action != StartupAction::ServiceDispatcher && parsed.action != action)
    {
        const std::wstring_view previous = parsed.actionOption.empty() ? GetActionOptionName(parsed.action) : std::wstring_view(parsed.actionOption);
        parsed.errorMessage              = std::format(L"Options '{}' and '{}' cannot be used together.", previous, option);
        return false;
    }

    parsed.action       = action;
    parsed.actionOption = std::wstring(option);
    return true;
}

[[nodiscard]] bool TryWriteTextToHandle(HANDLE handle, std::wstring_view text) noexcept
{
    if (! handle || handle == INVALID_HANDLE_VALUE)
    {
        return false;
    }

    DWORD mode = 0u;
    if (::GetConsoleMode(handle, &mode) != FALSE)
    {
        DWORD written = 0u;
        return ::WriteConsoleW(handle, text.data(), static_cast<DWORD>(text.size()), &written, nullptr) != FALSE;
    }

    const int bytesNeeded = ::WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (bytesNeeded <= 0)
    {
        return false;
    }

    std::string utf8(static_cast<size_t>(bytesNeeded), '\0');
    const int converted = ::WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), utf8.data(), bytesNeeded, nullptr, nullptr);
    if (converted != bytesNeeded)
    {
        return false;
    }

    DWORD written = 0u;
    return ::WriteFile(handle, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr) != FALSE;
}

void WriteConsoleText(std::wstring_view text, bool errorOutput) noexcept
{
    HANDLE stdHandle = ::GetStdHandle(errorOutput ? STD_ERROR_HANDLE : STD_OUTPUT_HANDLE);
    if (TryWriteTextToHandle(stdHandle, text))
    {
        return;
    }

    HANDLE alternateHandle = ::GetStdHandle(errorOutput ? STD_OUTPUT_HANDLE : STD_ERROR_HANDLE);
    if (TryWriteTextToHandle(alternateHandle, text))
    {
        return;
    }

    const bool hadConsole = ::GetConsoleWindow() != nullptr;
    if (! hadConsole)
    {
        if (::AttachConsole(ATTACH_PARENT_PROCESS) == FALSE)
        {
            const DWORD error = ::GetLastError();
            if (error != ERROR_ACCESS_DENIED)
            {
                static_cast<void>(::AllocConsole());
            }
        }
    }

    wil::unique_handle console(::CreateFileW(L"CONOUT$", GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0u, nullptr));
    if (console && TryWriteTextToHandle(console.get(), text))
    {
        return;
    }

    std::wstring boxed(text);
    const wchar_t* caption = errorOutput ? L"RedSalamander Search Service Error" : L"RedSalamander Search Service";
    const UINT iconFlags   = errorOutput ? MB_ICONERROR : MB_ICONINFORMATION;
    static_cast<void>(::MessageBoxW(nullptr, boxed.c_str(), caption, MB_OK | iconFlags));
}

[[nodiscard]] std::wstring GetExecutableLeafName() noexcept
{
    wchar_t buffer[MAX_PATH] = {};
    const DWORD written      = ::GetModuleFileNameW(nullptr, buffer, static_cast<DWORD>(std::size(buffer)));
    if (written == 0u || written >= std::size(buffer))
    {
        return L"RedSalamanderSearchService.exe";
    }

    std::wstring_view path(buffer, written);
    const size_t separator = path.find_last_of(L"\\/");
    return separator == std::wstring_view::npos ? std::wstring(path) : std::wstring(path.substr(separator + 1u));
}

[[nodiscard]] std::wstring GetServiceDisplayName() noexcept
{
#ifdef _DEBUG
    return L"RedSalamander Search Service (Debug)";
#else
    return L"RedSalamander Search Service";
#endif
}

[[nodiscard]] std::wstring GetServiceDescription() noexcept
{
#ifdef _DEBUG
    return L"Indexes local file metadata for RedSalamander instant search (Debug build).";
#else
    return L"Indexes local file metadata for RedSalamander instant search.";
#endif
}

[[nodiscard]] bool IsConsoleHandle(HANDLE handle) noexcept
{
    if (! handle || handle == INVALID_HANDLE_VALUE)
    {
        return false;
    }

    DWORD mode = 0u;
    return ::GetConsoleMode(handle, &mode) != FALSE;
}

[[nodiscard]] bool HasUsableStdHandle(HANDLE handle) noexcept
{
    if (! handle || handle == INVALID_HANDLE_VALUE)
    {
        return false;
    }

    ::SetLastError(NO_ERROR);
    const DWORD fileType = ::GetFileType(handle);
    return fileType != FILE_TYPE_UNKNOWN || ::GetLastError() == NO_ERROR;
}

template <size_t N> void CopyFixedText(std::array<wchar_t, N>& destination, const wchar_t* source) noexcept
{
    destination.fill(L'\0');
    if (source == nullptr || N == 0u)
    {
        return;
    }

    size_t index = 0u;
    for (; index + 1u < N && source[index] != L'\0'; ++index)
    {
        destination[index] = source[index];
    }
    destination[index] = L'\0';
}

template <size_t N> [[nodiscard]] std::wstring_view ViewFixedText(const std::array<wchar_t, N>& source) noexcept
{
    const size_t length = ::wcsnlen(source.data(), N);
    return std::wstring_view(source.data(), length);
}

[[nodiscard]] ForegroundDashboardSnapshot CaptureForegroundDashboardSnapshot(const ForegroundConsoleState& state) noexcept
{
    ForegroundDashboardSnapshot snapshot{};
    snapshot.eventType                  = static_cast<SearchServiceBroker::ServerEventType>(state.lastEvent.load(std::memory_order_relaxed));
    snapshot.lastActivityEvent          = static_cast<SearchServiceBroker::ServerEventType>(state.lastActivityEvent.load(std::memory_order_relaxed));
    snapshot.lastHr                     = static_cast<HRESULT>(state.lastHr.load(std::memory_order_relaxed));
    snapshot.lastActivityHr             = static_cast<HRESULT>(state.lastActivityHr.load(std::memory_order_relaxed));
    snapshot.handledRequests            = state.handledRequests.load(std::memory_order_relaxed);
    snapshot.totalConnections           = state.totalConnections.load(std::memory_order_relaxed);
    snapshot.lastActivityTickMs         = state.lastActivityTickMs.load(std::memory_order_relaxed);
    snapshot.activeRequestStartedTickMs = state.activeRequestStartedTickMs.load(std::memory_order_relaxed);
    snapshot.uiRevision                 = state.uiRevision.load(std::memory_order_relaxed);
    snapshot.stopRequested              = state.stopRequested.load(std::memory_order_relaxed);

    {
        std::scoped_lock lock(state.detailsMutex);
        snapshot.requestSnapshot.requestType            = state.requestType;
        snapshot.requestSnapshot.requestPhase           = state.requestPhase;
        snapshot.requestSnapshot.warningFlags           = state.warningFlags;
        snapshot.requestSnapshot.requestFlags           = state.requestFlags;
        snapshot.requestSnapshot.requestNameMode        = state.requestNameMode;
        snapshot.requestSnapshot.storeState             = static_cast<LocalSearchIndexCore::StoreState>(state.storeState);
        snapshot.requestSnapshot.syncPhase              = static_cast<LocalSearchIndexCore::SyncPhase>(state.syncPhase);
        snapshot.requestSnapshot.queryExecutionMode     = static_cast<LocalSearchIndexCore::QueryExecutionMode>(state.queryExecutionMode);
        snapshot.requestSnapshot.fallbackReason         = static_cast<LocalSearchIndexCore::FallbackReason>(state.fallbackReason);
        snapshot.requestSnapshot.completedRoots         = state.completedRoots;
        snapshot.requestSnapshot.totalRoots             = state.totalRoots;
        snapshot.requestSnapshot.batchesSent            = state.batchesSent;
        snapshot.requestSnapshot.scannedDirectories     = state.scannedDirectories;
        snapshot.requestSnapshot.scannedFiles           = state.scannedFiles;
        snapshot.requestSnapshot.candidateFiles         = state.candidateFiles;
        snapshot.requestSnapshot.matchedEntries         = state.matchedEntries;
        snapshot.requestSnapshot.snapshotFileBytes      = state.snapshotFileBytes;
        snapshot.requestSnapshot.estimatedMemoryBytes   = state.estimatedMemoryBytes;
        snapshot.requestSnapshot.ensureReadyDurationMs  = state.ensureReadyDurationMs;
        snapshot.requestSnapshot.executeQueryDurationMs = state.executeQueryDurationMs;
        snapshot.requestSnapshot.rootPath               = std::wstring(ViewFixedText(state.rootPath));
        snapshot.requestSnapshot.namePattern            = std::wstring(ViewFixedText(state.namePattern));
        snapshot.requestSnapshot.activeRoot             = std::wstring(ViewFixedText(state.activeRoot));
        snapshot.requestSnapshot.currentPath            = std::wstring(ViewFixedText(state.currentPath));
        snapshot.eventHistory.assign(state.pendingEvents.begin(), state.pendingEvents.end());
    }

    return snapshot;
}

[[nodiscard]] std::wstring FormatElapsedTime(std::chrono::steady_clock::duration elapsed) noexcept
{
    const auto totalSeconds = std::chrono::duration_cast<std::chrono::seconds>(elapsed).count();
    const auto hours        = totalSeconds / 3600;
    const auto minutes      = (totalSeconds % 3600) / 60;
    const auto seconds      = totalSeconds % 60;
    return std::format(
        L"{:02}:{:02}:{:02}", static_cast<unsigned long long>(hours), static_cast<unsigned long long>(minutes), static_cast<unsigned long long>(seconds));
}

[[nodiscard]] std::string Utf8FromWide(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    if (text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int sourceLength = static_cast<int>(text.size());
    int required           = ::WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), sourceLength, nullptr, 0, nullptr, nullptr);
    DWORD flags            = WC_ERR_INVALID_CHARS;
    if (required == 0)
    {
        flags    = 0u;
        required = ::WideCharToMultiByte(CP_UTF8, flags, text.data(), sourceLength, nullptr, 0, nullptr, nullptr);
    }
    if (required <= 0)
    {
        return {};
    }

    std::string utf8(static_cast<size_t>(required), '\0');
    const int written = ::WideCharToMultiByte(CP_UTF8, flags, text.data(), sourceLength, utf8.data(), required, nullptr, nullptr);
    if (written <= 0)
    {
        return {};
    }

    utf8.resize(static_cast<size_t>(written));
    return utf8;
}

[[nodiscard]] std::string Utf8FromWidePreservingEmpty(std::wstring_view text, std::string_view fallback = "-") noexcept
{
    if (text.empty())
    {
        return std::string(fallback);
    }

    std::string utf8 = Utf8FromWide(text);
    if (utf8.empty())
    {
        return std::string(fallback);
    }

    return utf8;
}

[[nodiscard]] std::wstring_view GetForegroundEventText(SearchServiceBroker::ServerEventType eventType) noexcept
{
    switch (eventType)
    {
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING: return L"Starting";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_WAITING_FOR_CLIENT: return L"Waiting for client";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_CLIENT_CONNECTED: return L"Client connected";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_REQUEST_RECEIVED: return L"Request received";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_QUERY_PROGRESS: return L"Query progress";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_QUERY_BATCH_SENT: return L"Query batch sent";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_QUERY_COMPLETED: return L"Query completed";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_PROGRESS: return L"Startup warmup";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_FAILED: return L"Startup warmup failed";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_COMPLETED: return L"Startup warmup done";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_QUEUED: return L"Maintenance queued";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_RUNNING: return L"Maintenance running";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_COMPLETED: return L"Maintenance completed";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_REQUEST_HANDLED: return L"Request handled";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STOPPING: return L"Stopping";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STOPPED: return L"Stopped";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_ERROR: return L"Error";
        default: return L"Unknown";
    }
}

[[nodiscard]] uint32_t GetForegroundEventBit(SearchServiceBroker::ServerEventType eventType) noexcept
{
    const uint32_t eventIndex = static_cast<uint32_t>(eventType);
    return eventIndex < 32u ? (1u << eventIndex) : 0u;
}

[[nodiscard]] std::wstring_view GetRequestTypeText(const SearchServiceBroker::ServerRequestType requestType) noexcept
{
    switch (requestType)
    {
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_STATUS: return L"Status";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_QUERY: return L"Query";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_REBUILD: return L"Rebuild";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_COMPACT: return L"Compact";
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE:
        default: return L"None";
    }
}

[[nodiscard]] std::wstring_view GetSearchPhaseText(const FileSystemSearchPhase phase) noexcept
{
    switch (phase)
    {
        case FILESYSTEM_SEARCH_PHASE_INITIALIZING: return L"Initializing";
        case FILESYSTEM_SEARCH_PHASE_ENUMERATING: return L"Enumerating";
        case FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP: return L"Index lookup";
        case FILESYSTEM_SEARCH_PHASE_CONTENT_SCAN: return L"Content scan";
        case FILESYSTEM_SEARCH_PHASE_COMPLETED: return L"Completed";
        default: return L"Unknown";
    }
}

[[nodiscard]] std::wstring BuildForegroundStoreStateText(const ForegroundRequestSnapshot& snapshot) noexcept
{
    return std::wstring(LocalSearchIndexCore::GetStoreStateText(snapshot.storeState));
}

[[nodiscard]] std::wstring BuildForegroundSyncStateText(const ForegroundRequestSnapshot& snapshot) noexcept
{
    std::wstring text(LocalSearchIndexCore::GetSyncPhaseText(snapshot.syncPhase));
    if (snapshot.totalRoots != 0u)
    {
        text.append(std::format(L" {}/{}", snapshot.completedRoots, snapshot.totalRoots));
    }
    if (! snapshot.activeRoot.empty())
    {
        text.append(L" @ ");
        text.append(snapshot.activeRoot);
    }
    return text;
}

[[nodiscard]] std::wstring BuildForegroundExecutionModeText(const ForegroundRequestSnapshot& snapshot) noexcept
{
    std::wstring text(LocalSearchIndexCore::GetQueryExecutionModeText(snapshot.queryExecutionMode));
    if (snapshot.fallbackReason != LocalSearchIndexCore::FallbackReason::None)
    {
        text.append(std::format(L" ({})", LocalSearchIndexCore::GetFallbackReasonText(snapshot.fallbackReason)));
    }
    if (snapshot.warningFlags != FILESYSTEM_SEARCH_WARNING_NONE)
    {
        text.append(std::format(L" | warnings=0x{:08X}", static_cast<unsigned long>(snapshot.warningFlags)));
    }
    return text;
}

[[nodiscard]] std::wstring BuildForegroundStorePathText(const LocalSearchIndexCore::PersistentStoreInfo& storeInfo);
[[nodiscard]] std::wstring BuildForegroundMaintenanceStateText(const LocalSearchIndexCore::PersistentStoreInfo& storeInfo,
                                                               bool maintenanceQueued,
                                                               bool maintenanceRunning);
[[nodiscard]] std::wstring BuildForegroundReadinessText(const LocalSearchIndexCore::PersistentStoreInfo& storeInfo);

[[nodiscard]] std::wstring FormatCompactBytes(const uint64_t bytes) noexcept
{
    if (bytes >= (1024ull * 1024ull))
    {
        return std::format(L"{:.1f} MB", static_cast<double>(bytes) / (1024.0 * 1024.0));
    }
    if (bytes >= 1024ull)
    {
        return std::format(L"{:.1f} KB", static_cast<double>(bytes) / 1024.0);
    }

    return std::format(L"{} B", bytes);
}

[[nodiscard]] std::wstring TruncateMiddle(std::wstring_view text, size_t maxCharacters) noexcept
{
    if (text.size() <= maxCharacters || maxCharacters < 8u)
    {
        return std::wstring(text);
    }

    const size_t prefixLength = (maxCharacters - 3u) / 2u;
    const size_t suffixLength = maxCharacters - 3u - prefixLength;

    std::wstring result;
    result.reserve(maxCharacters);
    result.append(text.substr(0u, prefixLength));
    result.append(L"...");
    result.append(text.substr(text.size() - suffixLength));
    return result;
}

[[nodiscard]] std::wstring FormatQueryModeSummary(const uint32_t requestNameMode, const uint32_t requestFlags) noexcept
{
    std::wstring summary;

    switch (static_cast<FileSystemSearchNameMode>(requestNameMode))
    {
        case FILESYSTEM_SEARCH_NAME_WILDCARD: summary = L"Wildcard"; break;
        case FILESYSTEM_SEARCH_NAME_LITERAL: summary = L"Literal"; break;
        case FILESYSTEM_SEARCH_NAME_REGEX: summary = L"Regex"; break;
        case FILESYSTEM_SEARCH_NAME_DISABLED:
        default: summary = L"Disabled"; break;
    }

    summary.append(L" | ");
    summary.append((requestFlags & FILESYSTEM_SEARCH_RECURSIVE) != 0u ? L"recursive" : L"single level");
    summary.append(L" | ");

    if ((requestFlags & FILESYSTEM_SEARCH_INCLUDE_FILES) != 0u && (requestFlags & FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES) != 0u)
    {
        summary.append(L"files+dirs");
    }
    else if ((requestFlags & FILESYSTEM_SEARCH_INCLUDE_FILES) != 0u)
    {
        summary.append(L"files");
    }
    else if ((requestFlags & FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES) != 0u)
    {
        summary.append(L"dirs");
    }
    else
    {
        summary.append(L"none");
    }

    if ((requestFlags & FILESYSTEM_SEARCH_MATCH_CASE_NAME) != 0u)
    {
        summary.append(L" | match-case");
    }
    return summary;
}

[[nodiscard]] std::wstring_view GetBuildFlavorText() noexcept
{
#if defined(__SANITIZE_ADDRESS__)
    return L"ASan Debug";
#elif defined(_DEBUG)
    return L"Debug";
#else
    return L"Release";
#endif
}

constexpr std::wstring_view kAnsiReset                 = L"\x1b[0m";
constexpr std::wstring_view kAnsiBold                  = L"\x1b[1m";
constexpr std::wstring_view kAnsiDim                   = L"\x1b[38;5;250m";
constexpr std::wstring_view kAnsiInfo                  = L"\x1b[38;5;81m";
constexpr std::wstring_view kAnsiAccent                = L"\x1b[1;38;5;82m";
constexpr std::wstring_view kAnsiWarm                  = L"\x1b[1;38;5;214m";
constexpr uint64_t kForegroundStallThresholdMs         = 5'000u;
constexpr std::wstring_view kForegroundSpinnerFrames[] = {L"|", L"/", L"-", L"\\"};

[[nodiscard]] bool TryEnableVirtualTerminal(HANDLE handle, DWORD& outOriginalMode) noexcept
{
    outOriginalMode = 0u;
    if (! IsConsoleHandle(handle))
    {
        return false;
    }

    DWORD mode = 0u;
    if (::GetConsoleMode(handle, &mode) == FALSE)
    {
        return false;
    }

    outOriginalMode         = mode;
    const DWORD desiredMode = mode | ENABLE_PROCESSED_OUTPUT | ENABLE_VIRTUAL_TERMINAL_PROCESSING;
    if (::SetConsoleMode(handle, desiredMode) == FALSE)
    {
        return false;
    }

    return true;
}

void AppendStyledText(std::wstring& out, std::wstring_view text, std::wstring_view style, bool useAnsiColors)
{
    if (useAnsiColors && ! style.empty())
    {
        out.append(style);
    }
    out.append(text);
    if (useAnsiColors && ! style.empty())
    {
        out.append(kAnsiReset);
    }
}

void AppendSeparator(std::wstring& out, bool useAnsiColors)
{
    AppendStyledText(out, std::wstring(80u, L'='), kAnsiAccent, useAnsiColors);
    out.append(L"\r\n");
}
[[nodiscard]] std::wstring_view GetForegroundEventText(SearchServiceBroker::ServerEventType eventType, const HRESULT hr) noexcept
{
    if (FAILED(hr) && eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_QUERY_COMPLETED)
    {
        return L"Query failed";
    }
    if (FAILED(hr) && eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_REQUEST_HANDLED)
    {
        return L"Request failed";
    }

    return GetForegroundEventText(eventType);
}
[[nodiscard]] std::wstring FormatActivityAge(uint64_t tickMs) noexcept
{
    if (tickMs == 0u)
    {
        return L"No client activity yet";
    }

    const ULONGLONG now       = ::GetTickCount64();
    const ULONGLONG elapsedMs = now >= tickMs ? (now - tickMs) : 0u;
    return std::format(L"{} ago", FormatElapsedTime(std::chrono::milliseconds(elapsedMs)));
}

[[nodiscard]] uint64_t GetActivityElapsedMs(uint64_t tickMs) noexcept
{
    if (tickMs == 0u)
    {
        return 0u;
    }

    const ULONGLONG now = ::GetTickCount64();
    return now >= tickMs ? static_cast<uint64_t>(now - tickMs) : 0u;
}

[[nodiscard]] std::wstring_view GetForegroundSpinnerFrame() noexcept
{
    constexpr size_t kFrameCount = sizeof(kForegroundSpinnerFrames) / sizeof(kForegroundSpinnerFrames[0]);
    const size_t frameIndex      = static_cast<size_t>((::GetTickCount64() / 200u) % kFrameCount);
    return kForegroundSpinnerFrames[frameIndex];
}

[[nodiscard]] std::wstring BuildForegroundPulseBar(size_t width) noexcept
{
    const size_t innerWidth = std::clamp(width, static_cast<size_t>(10u), static_cast<size_t>(20u));
    if (innerWidth <= 2u)
    {
        return L"[]";
    }

    const size_t segmentWidth = std::min(static_cast<size_t>(4u), innerWidth);
    const size_t travel       = innerWidth > segmentWidth ? (innerWidth - segmentWidth) : 0u;
    const size_t cycleLength  = travel == 0u ? 1u : (travel * 2u);
    const size_t cycleIndex   = static_cast<size_t>((::GetTickCount64() / 160u) % cycleLength);
    const size_t position     = travel == 0u ? 0u : (cycleIndex <= travel ? cycleIndex : (cycleLength - cycleIndex));

    std::wstring bar;
    bar.reserve(innerWidth + 2u);
    bar.push_back(L'[');
    for (size_t index = 0u; index < innerWidth; ++index)
    {
        const bool active = index >= position && index < (position + segmentWidth);
        bar.push_back(active ? L'=' : L' ');
    }
    bar.push_back(L']');
    return bar;
}

[[nodiscard]] bool IsForegroundRequestInFlight(const ForegroundRequestSnapshot& requestSnapshot) noexcept
{
    return requestSnapshot.requestType != SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE &&
           requestSnapshot.requestPhase != FILESYSTEM_SEARCH_PHASE_COMPLETED;
}

[[nodiscard]] bool IsForegroundRequestStalled(const ForegroundRequestSnapshot& requestSnapshot, uint64_t lastActivityTickMs, bool stopRequested) noexcept
{
    return ! stopRequested && IsForegroundRequestInFlight(requestSnapshot) && GetActivityElapsedMs(lastActivityTickMs) >= kForegroundStallThresholdMs;
}

[[nodiscard]] bool ShouldForegroundDashboardAnimate(const ForegroundConsoleState& state) noexcept
{
    return state.stopRequested.load(std::memory_order_relaxed) || state.activeRequestStartedTickMs.load(std::memory_order_relaxed) != 0u;
}

void AppendDashboardField(std::wstring& out, std::wstring_view label, std::wstring_view value, bool useAnsiColors, std::wstring_view valueStyle = {}) noexcept
{
    AppendStyledText(out, label, kAnsiDim, useAnsiColors);
    if (label.size() < 9u)
    {
        out.append(9u - label.size(), L' ');
    }
    out.append(L": ");
    AppendStyledText(out, value, valueStyle, useAnsiColors);
    out.append(L"\r\n");
}

[[nodiscard]] std::wstring BuildForegroundDashboardActionDetails(SearchServiceBroker::ServerEventType eventType,
                                                                 const ForegroundRequestSnapshot& requestSnapshot,
                                                                 HRESULT hr) noexcept
{
    if (requestSnapshot.requestType == SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE)
    {
        return eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_CLIENT_CONNECTED ? L"Awaiting request header" : L"-";
    }

    std::wstring details;
    if (requestSnapshot.requestType == SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_QUERY)
    {
        details.assign(GetSearchPhaseText(requestSnapshot.requestPhase));

        if (! requestSnapshot.namePattern.empty())
        {
            details.append(L" | ");
            details.append(TruncateMiddle(requestSnapshot.namePattern, 28u));
        }

        if (! requestSnapshot.currentPath.empty())
        {
            details.append(L" | ");
            details.append(TruncateMiddle(requestSnapshot.currentPath, 56u));
        }
        else if (! requestSnapshot.rootPath.empty())
        {
            details.append(L" | ");
            details.append(TruncateMiddle(requestSnapshot.rootPath, 56u));
        }

        details.append(std::format(L" | d:{} f:{} c:{} m:{} b:{}",
                                   requestSnapshot.scannedDirectories,
                                   requestSnapshot.scannedFiles,
                                   requestSnapshot.candidateFiles,
                                   requestSnapshot.matchedEntries,
                                   requestSnapshot.batchesSent));

        if (eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_QUERY_COMPLETED ||
            eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_REQUEST_HANDLED || FAILED(hr))
        {
            details.append(std::format(L" | ready={} ms query={} ms store={} mem={}",
                                       requestSnapshot.ensureReadyDurationMs,
                                       requestSnapshot.executeQueryDurationMs,
                                       FormatCompactBytes(requestSnapshot.snapshotFileBytes),
                                       FormatCompactBytes(requestSnapshot.estimatedMemoryBytes)));
        }

        return details;
    }

    details.assign(GetRequestTypeText(requestSnapshot.requestType));
    if (! requestSnapshot.rootPath.empty())
    {
        details.append(L" | ");
        details.append(TruncateMiddle(requestSnapshot.rootPath, 72u));
    }
    return details;
}

[[nodiscard]] std::wstring BuildForegroundDashboardEventSummary(const ForegroundEventRecord& record) noexcept
{
    std::wstring summary = std::format(L"{} | {}", GetForegroundEventText(record.type, record.hr), FormatActivityAge(record.tickMs));

    const std::wstring details = BuildForegroundDashboardActionDetails(record.type, record.requestSnapshot, record.hr);
    if (! details.empty() && details != L"-")
    {
        summary.append(L" | ");
        summary.append(details);
    }

    if (FAILED(record.hr))
    {
        summary.append(std::format(L" | hr=0x{:08X}", static_cast<unsigned long>(record.hr)));
    }

    return summary;
}

[[nodiscard]] const ForegroundEventRecord* GetForegroundLatestMeaningfulEvent(const std::vector<ForegroundEventRecord>& history) noexcept
{
    for (auto it = history.rbegin(); it != history.rend(); ++it)
    {
        if (it->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING)
        {
            continue;
        }

        return &(*it);
    }

    return nullptr;
}

[[nodiscard]] std::wstring BuildForegroundHealthSummary(const ForegroundDashboardSnapshot& snapshot) noexcept
{
    const bool requestInFlight = IsForegroundRequestInFlight(snapshot.requestSnapshot);
    const bool requestStalled  = IsForegroundRequestStalled(snapshot.requestSnapshot, snapshot.lastActivityTickMs, snapshot.stopRequested);

    if (snapshot.stopRequested)
    {
        return std::format(L"Stopping | clients={} | handled={}", snapshot.totalConnections, snapshot.handledRequests);
    }
    if (requestStalled)
    {
        return std::format(L"Stalled {} | no progress {} | clients={} | handled={}",
                           GetRequestTypeText(snapshot.requestSnapshot.requestType),
                           FormatActivityAge(snapshot.lastActivityTickMs),
                           snapshot.totalConnections,
                           snapshot.handledRequests);
    }
    if (requestInFlight)
    {
        return std::format(L"Active {} | last progress {} | clients={} | handled={}",
                           GetRequestTypeText(snapshot.requestSnapshot.requestType),
                           FormatActivityAge(snapshot.lastActivityTickMs),
                           snapshot.totalConnections,
                           snapshot.handledRequests);
    }
    if (snapshot.eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_WAITING_FOR_CLIENT)
    {
        return std::format(L"Idle | clients={} | handled={}", snapshot.totalConnections, snapshot.handledRequests);
    }

    return std::format(L"Ready | last activity {} | clients={} | handled={}",
                       FormatActivityAge(snapshot.lastActivityTickMs),
                       snapshot.totalConnections,
                       snapshot.handledRequests);
}

[[nodiscard]] std::wstring BuildForegroundPulseSummary(const ForegroundDashboardSnapshot& snapshot) noexcept
{
    const bool requestInFlight  = IsForegroundRequestInFlight(snapshot.requestSnapshot);
    const bool requestStalled   = IsForegroundRequestStalled(snapshot.requestSnapshot, snapshot.lastActivityTickMs, snapshot.stopRequested);
    const uint64_t requestAgeMs = requestInFlight ? GetActivityElapsedMs(snapshot.activeRequestStartedTickMs) : 0u;

    if (snapshot.stopRequested)
    {
        return L"! shutdown requested";
    }
    if (requestStalled)
    {
        std::wstring summary = std::format(L"! stalled {} in {} | last progress {}",
                                           FormatElapsedTime(std::chrono::milliseconds(requestAgeMs)),
                                           GetSearchPhaseText(snapshot.requestSnapshot.requestPhase),
                                           FormatActivityAge(snapshot.lastActivityTickMs));
        if (! snapshot.requestSnapshot.currentPath.empty())
        {
            summary.append(L" | ");
            summary.append(TruncateMiddle(snapshot.requestSnapshot.currentPath, 64u));
        }
        return summary;
    }
    if (requestInFlight)
    {
        std::wstring summary = std::format(L"{} {} {} | {} | running {} | last progress {}",
                                           GetForegroundSpinnerFrame(),
                                           BuildForegroundPulseBar(14u),
                                           GetRequestTypeText(snapshot.requestSnapshot.requestType),
                                           GetSearchPhaseText(snapshot.requestSnapshot.requestPhase),
                                           FormatElapsedTime(std::chrono::milliseconds(requestAgeMs)),
                                           FormatActivityAge(snapshot.lastActivityTickMs));
        if (! snapshot.requestSnapshot.currentPath.empty())
        {
            summary.append(L" | ");
            summary.append(TruncateMiddle(snapshot.requestSnapshot.currentPath, 64u));
        }
        return summary;
    }
    if (snapshot.lastActivityTickMs != 0u)
    {
        return std::format(L". idle | last activity {}", FormatActivityAge(snapshot.lastActivityTickMs));
    }

    return L". waiting for first client";
}

[[nodiscard]] std::wstring BuildForegroundLastActionText(const ForegroundDashboardSnapshot& snapshot) noexcept
{
    if (const ForegroundEventRecord* record = GetForegroundLatestMeaningfulEvent(snapshot.eventHistory))
    {
        return std::format(L"{} | {}{}",
                           GetForegroundEventText(record->type, record->hr),
                           FormatActivityAge(record->tickMs),
                           FAILED(record->hr) ? std::format(L" | hr=0x{:08X}", static_cast<unsigned long>(record->hr)) : std::wstring{});
    }

    if (snapshot.lastActivityTickMs == 0u)
    {
        return L"No client activity yet";
    }

    return std::format(L"{} | {}{}",
                       GetForegroundEventText(snapshot.lastActivityEvent, snapshot.lastActivityHr),
                       FormatActivityAge(snapshot.lastActivityTickMs),
                       FAILED(snapshot.lastActivityHr) ? std::format(L" | hr=0x{:08X}", static_cast<unsigned long>(snapshot.lastActivityHr)) : std::wstring{});
}

[[nodiscard]] std::wstring BuildForegroundLastActionDetails(const ForegroundDashboardSnapshot& snapshot) noexcept
{
    if (const ForegroundEventRecord* record = GetForegroundLatestMeaningfulEvent(snapshot.eventHistory))
    {
        return BuildForegroundDashboardActionDetails(record->type, record->requestSnapshot, record->hr);
    }

    if (snapshot.lastActivityTickMs != 0u)
    {
        return BuildForegroundDashboardActionDetails(snapshot.lastActivityEvent, snapshot.requestSnapshot, snapshot.lastActivityHr);
    }

    return L"-";
}

[[nodiscard]] std::wstring BuildForegroundRequestTypeSummary(const ForegroundDashboardSnapshot& snapshot) noexcept
{
    if (snapshot.requestSnapshot.requestType == SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE)
    {
        return snapshot.eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_CLIENT_CONNECTED ? L"Connection open, awaiting request header"
                                                                                                       : L"No request received yet";
    }

    std::wstring summary(GetRequestTypeText(snapshot.requestSnapshot.requestType));
    if (snapshot.requestSnapshot.requestType == SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_QUERY)
    {
        summary.append(L" | ");
        summary.append(FormatQueryModeSummary(snapshot.requestSnapshot.requestNameMode, snapshot.requestSnapshot.requestFlags));
        if (! snapshot.requestSnapshot.namePattern.empty())
        {
            summary.append(L" | ");
            summary.append(snapshot.requestSnapshot.namePattern);
        }
    }

    return summary;
}

[[nodiscard]] std::wstring BuildForegroundFocusSummary(const ForegroundDashboardSnapshot& snapshot) noexcept
{
    if (snapshot.requestSnapshot.requestType == SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE)
    {
        return snapshot.eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_CLIENT_CONNECTED ? L"Awaiting request header" : L"Idle";
    }

    std::wstring summary(GetSearchPhaseText(snapshot.requestSnapshot.requestPhase));
    if (! snapshot.requestSnapshot.currentPath.empty())
    {
        summary.append(L" | ");
        summary.append(snapshot.requestSnapshot.currentPath);
    }
    return summary;
}

[[nodiscard]] std::wstring BuildForegroundScanSummary(const ForegroundDashboardSnapshot& snapshot) noexcept
{
    if (snapshot.requestSnapshot.requestType != SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_QUERY)
    {
        return L"-";
    }

    return std::format(L"dirs={} | files={} | candidates={} | matches={} | batches={}",
                       snapshot.requestSnapshot.scannedDirectories,
                       snapshot.requestSnapshot.scannedFiles,
                       snapshot.requestSnapshot.candidateFiles,
                       snapshot.requestSnapshot.matchedEntries,
                       snapshot.requestSnapshot.batchesSent);
}

[[nodiscard]] std::wstring BuildForegroundResultSummary(const ForegroundDashboardSnapshot& snapshot) noexcept
{
    if (snapshot.requestSnapshot.requestType != SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_QUERY)
    {
        return L"-";
    }

    return std::format(L"store={} | memory={} | ready={} ms | query={} ms",
                       FormatCompactBytes(snapshot.requestSnapshot.snapshotFileBytes),
                       FormatCompactBytes(snapshot.requestSnapshot.estimatedMemoryBytes),
                       snapshot.requestSnapshot.ensureReadyDurationMs,
                       snapshot.requestSnapshot.executeQueryDurationMs);
}

[[nodiscard]] ftxui::Color GetDashboardToneColor(const ForegroundDashboardSnapshot& snapshot) noexcept
{
    if (FAILED(snapshot.lastHr) || FAILED(snapshot.lastActivityHr))
    {
        return ftxui::Color::RedLight;
    }
    if (snapshot.stopRequested || IsForegroundRequestStalled(snapshot.requestSnapshot, snapshot.lastActivityTickMs, snapshot.stopRequested))
    {
        return ftxui::Color::YellowLight;
    }
    if (IsForegroundRequestInFlight(snapshot.requestSnapshot))
    {
        return ftxui::Color::GreenLight;
    }
    if (snapshot.eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_WAITING_FOR_CLIENT)
    {
        return ftxui::Color::GrayLight;
    }
    return ftxui::Color::CyanLight;
}

[[nodiscard]] ftxui::Color GetDashboardEventColor(const ForegroundEventRecord& record) noexcept
{
    if (FAILED(record.hr))
    {
        return ftxui::Color::RedLight;
    }

    switch (record.type)
    {
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING:
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_REQUEST_RECEIVED: return ftxui::Color::BlueLight;
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_QUERY_COMPLETED:
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_COMPLETED:
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_COMPLETED:
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_REQUEST_HANDLED: return ftxui::Color::GreenLight;
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_QUERY_PROGRESS:
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_PROGRESS:
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_RUNNING: return ftxui::Color::CyanLight;
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_QUERY_BATCH_SENT:
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_CLIENT_CONNECTED:
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_QUEUED: return ftxui::Color::YellowLight;
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_FAILED:
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_ERROR: return ftxui::Color::RedLight;
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STOPPING:
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STOPPED: return ftxui::Color::YellowLight;
        case SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_WAITING_FOR_CLIENT: return ftxui::Color::GrayLight;
        default: return ftxui::Color::White;
    }
}

[[nodiscard]] ftxui::Color GetDashboardMutedColor() noexcept
{
    return ftxui::Color::GrayLight;
}

[[nodiscard]] ftxui::Color GetDashboardBadgeForegroundColor(ftxui::Color background) noexcept
{
    if (background == ftxui::Color::GreenLight || background == ftxui::Color::CyanLight || background == ftxui::Color::YellowLight ||
        background == ftxui::Color::White || background == ftxui::Color::GrayLight)
    {
        return ftxui::Color::Black;
    }

    return ftxui::Color::White;
}

[[nodiscard]] ftxui::Element MakeDashboardBadge(std::string_view label, std::string_view value, ftxui::Color color) noexcept
{
    return ftxui::text(std::format(" {} {} ", label, value)) | ftxui::color(GetDashboardBadgeForegroundColor(color)) | ftxui::bgcolor(color) | ftxui::bold;
}

[[nodiscard]] ftxui::Element MakeDashboardField(std::string_view label, std::wstring_view value, ftxui::Color color = ftxui::Color::White) noexcept
{
    return ftxui::hbox({
               ftxui::text(std::string(label)) | ftxui::color(GetDashboardMutedColor()) | ftxui::size(ftxui::WIDTH, ftxui::EQUAL, 12),
               ftxui::separatorEmpty(),
               ftxui::paragraph(Utf8FromWidePreservingEmpty(value)) | ftxui::color(color) | ftxui::flex,
           }) |
           ftxui::xflex;
}

[[nodiscard]] ftxui::Element MakeDashboardMetricCard(std::string_view label, std::wstring_view value, std::string_view subtitle, ftxui::Color color) noexcept
{
    return ftxui::window(ftxui::text(std::string(label)),
                         ftxui::vbox({
                             ftxui::text(Utf8FromWidePreservingEmpty(value)) | ftxui::color(color) | ftxui::bold,
                             ftxui::text(std::string(subtitle)) | ftxui::color(GetDashboardMutedColor()),
                         })) |
           ftxui::xflex;
}

[[nodiscard]] std::wstring GetForegroundRunModeText(uint32_t maxRequestsBeforeExit, uint32_t disconnectAfterBatches);

class ForegroundDashboardController final
{
public:
    ForegroundDashboardController(const ForegroundConsoleState& state,
                                  std::wstring pipeName,
                                  std::wstring storageRoot,
                                  LocalSearchIndexCore::PersistentStoreKind persistentStoreKind,
                                  std::wstring sqliteDatabasePath,
                                  uint32_t protocolVersion,
                                  uint32_t maxRequestsBeforeExit,
                                  uint32_t disconnectAfterBatches,
                                  std::chrono::steady_clock::time_point startTime) noexcept
        : _state(state),
          _pipeName(std::move(pipeName)),
          _storageRoot(std::move(storageRoot)),
          _persistentStoreKind(persistentStoreKind),
          _sqliteDatabasePath(std::move(sqliteDatabasePath)),
          _protocolVersion(protocolVersion),
          _maxRequestsBeforeExit(maxRequestsBeforeExit),
          _disconnectAfterBatches(disconnectAfterBatches),
          _startTime(startTime)
    {
    }

    ForegroundDashboardController(const ForegroundDashboardController&)            = delete;
    ForegroundDashboardController& operator=(const ForegroundDashboardController&) = delete;

    void Run(const std::atomic<bool>& serverExited) noexcept
    {
        using namespace ftxui;

        Refresh();

        ScreenInteractive screen = ScreenInteractive::FullscreenAlternateScreen();
        screen.ForceHandleCtrlC(false);

        auto pageToggle                  = Toggle(&_pageEntries, &_selectedPage);
        auto historyOptions              = MenuOption::VerticalAnimated();
        historyOptions.underline.enabled = false;
        historyOptions.entries_option.animated_colors.foreground.Set(Color::White, Color::Black);
        historyOptions.entries_option.animated_colors.background.Set(Color::Black, Color::CyanLight);
        auto historyMenu = Menu(&_historyLabels, &_selectedHistory, historyOptions);

        auto overview = Renderer([this] { return BuildOverviewPage(); });
        auto history  = Renderer(historyMenu, [this, &historyMenu] { return BuildHistoryPage(historyMenu->Render()); });
        auto pages    = Container::Tab({overview, history}, &_selectedPage);
        auto root     = Container::Vertical({pageToggle, pages});

        auto app = CatchEvent(root,
                              [this, &screen, &serverExited](Event event)
        {
            if (event == Event::Custom)
            {
                Refresh();
                if (serverExited.load(std::memory_order_acquire))
                {
                    screen.ExitLoopClosure()();
                }
                return true;
            }

            if (event == Event::Character('1'))
            {
                _selectedPage = 0;
                return true;
            }
            if (event == Event::Character('2'))
            {
                _selectedPage = 1;
                return true;
            }
            if (event == Event::Character('f') || event == Event::Character('F'))
            {
                _followLatest = ! _followLatest;
                if (_followLatest && ! _historyRecords.empty())
                {
                    _selectedHistory = 0;
                }
                return true;
            }
            if ((event == Event::ArrowUp || event == Event::ArrowDown || event == Event::PageUp || event == Event::PageDown || event == Event::Home ||
                 event == Event::End) &&
                _selectedPage == 1)
            {
                _followLatest = false;
            }

            return false;
        });

        auto document = Renderer(app,
                                 [this, &pageToggle, &pages]
        {
            return ftxui::vbox({
                       BuildHeader(),
                       ftxui::separator(),
                       pageToggle->Render(),
                       ftxui::separator(),
                       pages->Render() | ftxui::flex,
                       ftxui::separator(),
                       BuildFooter(),
                   }) |
                   ftxui::bgcolor(ftxui::Color::Black) | ftxui::color(ftxui::Color::White);
        });

        wil::unique_handle wakeEvent(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
        if (! wakeEvent)
        {
            return;
        }

        g_foregroundUiWakeEvent.store(wakeEvent.get(), std::memory_order_release);
        const auto restoreWakeEvent = wil::scope_exit([&] noexcept { g_foregroundUiWakeEvent.store(nullptr, std::memory_order_release); });

        std::jthread refreshThread([this, &screen, &serverExited, wake = wakeEvent.get()](std::stop_token stopToken) noexcept
        {
            while (! stopToken.stop_requested())
            {
                const DWORD waitMs     = ShouldForegroundDashboardAnimate(_state) ? 200u : INFINITE;
                const DWORD waitResult = ::WaitForSingleObject(wake, waitMs);
                if (waitResult == WAIT_OBJECT_0)
                {
                    static_cast<void>(::ResetEvent(wake));
                }

                screen.PostEvent(ftxui::Event::Custom);
                if (serverExited.load(std::memory_order_acquire))
                {
                    break;
                }
            }
        });

        screen.Loop(document);

        refreshThread.request_stop();
        static_cast<void>(::SetEvent(wakeEvent.get()));
    }

private:
    void Refresh() noexcept
    {
        _snapshot = CaptureForegroundDashboardSnapshot(_state);
        RebuildHistory();
    }

    void RebuildHistory() noexcept
    {
        const int previousSelection = _selectedHistory;
        _historyRecords.clear();
        _historyLabels.clear();

        for (auto it = _snapshot.eventHistory.rbegin(); it != _snapshot.eventHistory.rend(); ++it)
        {
            if (it->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING)
            {
                continue;
            }

            _historyRecords.push_back(*it);
            const std::wstring label = TruncateMiddle(
                std::format(L"{} | {}", GetForegroundEventText(it->type, it->hr), BuildForegroundDashboardActionDetails(it->type, it->requestSnapshot, it->hr)),
                96u);
            _historyLabels.push_back(Utf8FromWidePreservingEmpty(label, "<unavailable>"));
        }

        if (_historyRecords.empty())
        {
            _historyLabels.push_back("No activity yet");
        }

        if (_followLatest && ! _historyRecords.empty())
        {
            _selectedHistory = 0;
        }
        else if (_historyRecords.empty())
        {
            _selectedHistory = 0;
        }
        else
        {
            _selectedHistory = std::clamp(previousSelection, 0, static_cast<int>(_historyRecords.size() - 1u));
        }
    }

    [[nodiscard]] ftxui::Element BuildHeader() const
    {
        const ftxui::Color tone = GetDashboardToneColor(_snapshot);
        return ftxui::vbox({
            ftxui::hbox({
                ftxui::text("RedSalamander Search Service") | ftxui::color(ftxui::Color::White) | ftxui::bold,
                ftxui::filler(),
                MakeDashboardBadge("BUILD", Utf8FromWidePreservingEmpty(GetBuildFlavorText()).c_str(), ftxui::Color::GreenLight),
                ftxui::separatorEmpty(),
                MakeDashboardBadge("PID", std::format("{}", ::GetCurrentProcessId()), ftxui::Color::BlueLight),
                ftxui::separatorEmpty(),
                MakeDashboardBadge("UP", Utf8FromWidePreservingEmpty(FormatElapsedTime(std::chrono::steady_clock::now() - _startTime)).c_str(), tone),
            }),
            ftxui::hbox({
                ftxui::paragraph(Utf8FromWidePreservingEmpty(GetForegroundRunModeText(_maxRequestsBeforeExit, _disconnectAfterBatches))) |
                    ftxui::color(GetDashboardMutedColor()) | ftxui::flex,
                ftxui::separatorEmpty(),
                MakeDashboardBadge(_selectedPage == 0 ? "PAGE" : "PAGE",
                                   _selectedPage == 0 ? "Overview" : "History",
                                   _selectedPage == 0 ? ftxui::Color::CyanLight : ftxui::Color::YellowLight),
            }),
        });
    }

    [[nodiscard]] ftxui::Element BuildOverviewPage() const
    {
        const ftxui::Color tone                                   = GetDashboardToneColor(_snapshot);
        const std::wstring requestSummary                         = BuildForegroundRequestTypeSummary(_snapshot);
        const std::wstring focusSummary                           = BuildForegroundFocusSummary(_snapshot);
        const LocalSearchIndexCore::PersistentStoreInfo storeInfo = LocalSearchIndexCore::GetPersistentStoreInfo({
            .snapshotRootDirectory = _storageRoot,
            .persistentStoreKind   = _persistentStoreKind,
            .sqliteDatabasePath    = _sqliteDatabasePath,
            .sqliteAuthoritative   = _persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
        });
        const bool maintenanceQueued                              = _state.maintenanceQueued.load(std::memory_order_relaxed);
        const bool maintenanceRunning                             = _state.maintenanceRunning.load(std::memory_order_relaxed);
        const std::wstring syncSummary                            = BuildForegroundSyncStateText(_snapshot.requestSnapshot);
        const std::wstring executionSummary                       = BuildForegroundExecutionModeText(_snapshot.requestSnapshot);
        const std::wstring activeRoot =
            _snapshot.requestSnapshot.activeRoot.empty() ? _snapshot.requestSnapshot.rootPath : _snapshot.requestSnapshot.activeRoot;
        const std::wstring warmupCurrentRoot(ViewFixedText(_state.startupWarmupCurrentRoot));
        const std::wstring warmupSummary = std::format(L"{} | {}/{} failed={}{}",
                                                       _state.startupWarmupRunning ? L"running" : (_state.startupWarmupEnabled ? L"idle" : L"disabled"),
                                                       _state.startupWarmupCompletedRoots,
                                                       _state.startupWarmupTotalRoots,
                                                       _state.startupWarmupFailedRoots,
                                                       warmupCurrentRoot.empty() ? std::wstring() : std::format(L" @ {}", warmupCurrentRoot));

        return ftxui::vbox({
            ftxui::hbox({
                ftxui::window(ftxui::text("Live state"),
                              ftxui::vbox({
                                  MakeDashboardField("Health", BuildForegroundHealthSummary(_snapshot), tone),
                                  MakeDashboardField("Pulse", BuildForegroundPulseSummary(_snapshot), tone),
                                  MakeDashboardField("Last action", BuildForegroundLastActionText(_snapshot), ftxui::Color::CyanLight),
                                  MakeDashboardField("Action detail", BuildForegroundLastActionDetails(_snapshot), ftxui::Color::White),
                              })) |
                    ftxui::xflex,
                ftxui::window(ftxui::text("Session"),
                              ftxui::vbox({
                                  MakeDashboardField("Service", SearchServiceBroker::kServiceName, ftxui::Color::White),
                                  MakeDashboardField("Protocol", std::format(L"{}", _protocolVersion), ftxui::Color::BlueLight),
                                  MakeDashboardField("Pipe", TruncateMiddle(_pipeName, 72u), ftxui::Color::White),
                                  MakeDashboardField("Storage", TruncateMiddle(_storageRoot, 72u), ftxui::Color::White),
                              })) |
                    ftxui::xflex,
            }),
            ftxui::hbox({
                ftxui::window(ftxui::text("Current request"),
                              ftxui::vbox({
                                  MakeDashboardField("Type", requestSummary, ftxui::Color::GreenLight),
                                  MakeDashboardField("Root", TruncateMiddle(_snapshot.requestSnapshot.rootPath, 84u), ftxui::Color::White),
                                  MakeDashboardField("Focus", TruncateMiddle(focusSummary, 84u), ftxui::Color::CyanLight),
                                  MakeDashboardField("Pattern", TruncateMiddle(_snapshot.requestSnapshot.namePattern, 84u), ftxui::Color::YellowLight),
                              })) |
                    ftxui::xflex,
                ftxui::window(
                    ftxui::text("Search stats"),
                    ftxui::vbox({
                        ftxui::hbox({
                            MakeDashboardMetricCard(
                                "Dirs", std::format(L"{}", _snapshot.requestSnapshot.scannedDirectories), "enumerated", ftxui::Color::BlueLight),
                            MakeDashboardMetricCard("Files", std::format(L"{}", _snapshot.requestSnapshot.scannedFiles), "examined", ftxui::Color::BlueLight),
                            MakeDashboardMetricCard(
                                "Candidates", std::format(L"{}", _snapshot.requestSnapshot.candidateFiles), "selected", ftxui::Color::YellowLight),
                            MakeDashboardMetricCard(
                                "Matches", std::format(L"{}", _snapshot.requestSnapshot.matchedEntries), "emitted", ftxui::Color::GreenLight),
                        }),
                        ftxui::hbox({
                            MakeDashboardMetricCard("Batches", std::format(L"{}", _snapshot.requestSnapshot.batchesSent), "sent", ftxui::Color::YellowLight),
                            MakeDashboardMetricCard(
                                "Ready", std::format(L"{} ms", _snapshot.requestSnapshot.ensureReadyDurationMs), "warm-up", ftxui::Color::White),
                            MakeDashboardMetricCard(
                                "Query", std::format(L"{} ms", _snapshot.requestSnapshot.executeQueryDurationMs), "execution", ftxui::Color::White),
                            MakeDashboardMetricCard("Store", FormatCompactBytes(_snapshot.requestSnapshot.snapshotFileBytes), "disk", ftxui::Color::CyanLight),
                        }),
                        MakeDashboardField("Scan", BuildForegroundScanSummary(_snapshot), ftxui::Color::White),
                        MakeDashboardField("Result", BuildForegroundResultSummary(_snapshot), ftxui::Color::White),
                    })) |
                    ftxui::xflex,
            }),
            ftxui::hbox({
                ftxui::window(ftxui::text("Database"),
                              ftxui::vbox({
                                  MakeDashboardField("Backend", LocalSearchIndexCore::GetPersistentStoreKindText(storeInfo.kind), ftxui::Color::CyanLight),
                                  MakeDashboardField("Store", TruncateMiddle(BuildForegroundStorePathText(storeInfo), 96u), ftxui::Color::White),
                                  MakeDashboardField("Readiness",
                                                     BuildForegroundReadinessText(storeInfo),
                                                     storeInfo.readyForQueryCutover ? ftxui::Color::GreenLight : ftxui::Color::YellowLight),
                                  MakeDashboardField("Maintenance",
                                                     BuildForegroundMaintenanceStateText(storeInfo, maintenanceQueued, maintenanceRunning),
                                                     maintenanceRunning ? ftxui::Color::CyanLight : GetDashboardMutedColor()),
                                  MakeDashboardField("Runtime", BuildForegroundStoreStateText(_snapshot.requestSnapshot), ftxui::Color::YellowLight),
                              })) |
                    ftxui::xflex,
                ftxui::window(ftxui::text("Synchronization"),
                              ftxui::vbox({
                                  MakeDashboardField("Sync", syncSummary, ftxui::Color::CyanLight),
                                  MakeDashboardField("Warmup", warmupSummary, _state.startupWarmupRunning ? ftxui::Color::CyanLight : GetDashboardMutedColor()),
                                  MakeDashboardField("Search",
                                                     executionSummary,
                                                     _snapshot.requestSnapshot.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback
                                                         ? ftxui::Color::YellowLight
                                                         : ftxui::Color::GreenLight),
                                  MakeDashboardField("Active root", TruncateMiddle(activeRoot, 96u), ftxui::Color::White),
                              })) |
                    ftxui::xflex,
            }),
            ftxui::window(ftxui::text("Recent activity"), BuildRecentEventsList(8u)) | ftxui::flex,
        });
    }

    [[nodiscard]] ftxui::Element BuildHistoryPage(ftxui::Element historyMenuElement) const
    {
        return ftxui::hbox({
                   ftxui::window(
                       ftxui::text(_followLatest ? "History (follow latest)" : "History"),
                       ftxui::vbox({
                           ftxui::hbox({
                               MakeDashboardBadge("EVENTS", std::format("{}", _historyRecords.size()), ftxui::Color::BlueLight),
                               ftxui::separatorEmpty(),
                               MakeDashboardBadge("FOLLOW", _followLatest ? "ON" : "OFF", _followLatest ? ftxui::Color::GreenLight : GetDashboardMutedColor()),
                           }),
                           ftxui::separator(),
                           historyMenuElement | ftxui::frame | ftxui::vscroll_indicator | ftxui::flex,
                       })) |
                       ftxui::size(ftxui::WIDTH, ftxui::GREATER_THAN, 44),
                   ftxui::window(ftxui::text("Selected event"), BuildSelectedEventDetails()) | ftxui::xflex | ftxui::flex,
               }) |
               ftxui::flex;
    }

    [[nodiscard]] ftxui::Element BuildRecentEventsList(size_t limit) const
    {
        ftxui::Elements lines;
        size_t emitted = 0u;
        for (auto it = _snapshot.eventHistory.rbegin(); it != _snapshot.eventHistory.rend() && emitted < limit; ++it)
        {
            if (it->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING)
            {
                continue;
            }

            lines.push_back(ftxui::paragraph(Utf8FromWidePreservingEmpty(TruncateMiddle(BuildForegroundDashboardEventSummary(*it), 132u))) |
                            ftxui::color(GetDashboardEventColor(*it)));
            ++emitted;
        }

        if (lines.empty())
        {
            lines.push_back(ftxui::text("No completed client activity yet") | ftxui::color(GetDashboardMutedColor()));
        }

        return ftxui::vbox(std::move(lines));
    }

    [[nodiscard]] ftxui::Element BuildSelectedEventDetails() const
    {
        if (_historyRecords.empty())
        {
            return ftxui::vbox({
                ftxui::text("No event has been recorded yet.") | ftxui::color(GetDashboardMutedColor()),
            });
        }

        const ForegroundEventRecord& record = _historyRecords[static_cast<size_t>(_selectedHistory)];
        const std::wstring requestSummary   = record.requestSnapshot.requestType == SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE
                                                  ? L"None"
                                                  : std::wstring(GetRequestTypeText(record.requestSnapshot.requestType));

        return ftxui::vbox({
            MakeDashboardField("Event", GetForegroundEventText(record.type, record.hr), GetDashboardEventColor(record)),
            MakeDashboardField("Age", FormatActivityAge(record.tickMs), ftxui::Color::White),
            MakeDashboardField("HRESULT",
                               std::format(L"0x{:08X}", static_cast<unsigned long>(record.hr)),
                               FAILED(record.hr) ? ftxui::Color::RedLight : ftxui::Color::BlueLight),
            MakeDashboardField("Request", requestSummary, ftxui::Color::GreenLight),
            MakeDashboardField("Phase", GetSearchPhaseText(record.requestSnapshot.requestPhase), ftxui::Color::CyanLight),
            MakeDashboardField("Summary", BuildForegroundDashboardActionDetails(record.type, record.requestSnapshot, record.hr), ftxui::Color::White),
            MakeDashboardField("Root", TruncateMiddle(record.requestSnapshot.rootPath, 84u), ftxui::Color::White),
            MakeDashboardField("Path", TruncateMiddle(record.requestSnapshot.currentPath, 84u), ftxui::Color::White),
            MakeDashboardField(
                "Mode", FormatQueryModeSummary(record.requestSnapshot.requestNameMode, record.requestSnapshot.requestFlags), ftxui::Color::YellowLight),
            MakeDashboardField("Scan",
                               std::format(L"dirs={} | files={} | candidates={} | matches={} | batches={}",
                                           record.requestSnapshot.scannedDirectories,
                                           record.requestSnapshot.scannedFiles,
                                           record.requestSnapshot.candidateFiles,
                                           record.requestSnapshot.matchedEntries,
                                           record.requestSnapshot.batchesSent),
                               ftxui::Color::White),
            MakeDashboardField("Result",
                               std::format(L"store={} | memory={} | ready={} ms | query={} ms",
                                           FormatCompactBytes(record.requestSnapshot.snapshotFileBytes),
                                           FormatCompactBytes(record.requestSnapshot.estimatedMemoryBytes),
                                           record.requestSnapshot.ensureReadyDurationMs,
                                           record.requestSnapshot.executeQueryDurationMs),
                               ftxui::Color::White),
        });
    }

    [[nodiscard]] ftxui::Element BuildFooter() const
    {
        return ftxui::hbox({
            MakeDashboardBadge("1", "Overview", ftxui::Color::BlueLight),
            ftxui::separatorEmpty(),
            MakeDashboardBadge("2", "History", ftxui::Color::BlueLight),
            ftxui::separatorEmpty(),
            MakeDashboardBadge(
                "F", _followLatest ? "Following latest" : "Manual history", _followLatest ? ftxui::Color::GreenLight : ftxui::Color::YellowLight),
            ftxui::separatorEmpty(),
            MakeDashboardBadge("Arrows", "Browse history", GetDashboardMutedColor()),
            ftxui::separatorEmpty(),
            MakeDashboardBadge("Ctrl+C", "Stop service", ftxui::Color::RedLight),
        });
    }

    const ForegroundConsoleState& _state;
    std::wstring _pipeName;
    std::wstring _storageRoot;
    LocalSearchIndexCore::PersistentStoreKind _persistentStoreKind = LocalSearchIndexCore::PersistentStoreKind::Sqlite;
    std::wstring _sqliteDatabasePath;
    uint32_t _protocolVersion        = 0u;
    uint32_t _maxRequestsBeforeExit  = 0u;
    uint32_t _disconnectAfterBatches = 0u;
    std::chrono::steady_clock::time_point _startTime{};
    ForegroundDashboardSnapshot _snapshot{};
    std::vector<ForegroundEventRecord> _historyRecords;
    std::vector<std::string> _historyLabels;
    std::vector<std::string> _pageEntries{"Overview", "History"};
    int _selectedPage    = 0;
    int _selectedHistory = 0;
    bool _followLatest   = true;
};

[[nodiscard]] std::wstring BuildForegroundAsciiArt(bool useAnsiColors)
{
    std::wstring art;
    AppendStyledText(art, LR"(      _/\__      )", kAnsiWarm, useAnsiColors);
    AppendStyledText(art, L"RedSalamander Search Service", std::format(L"{}{}", kAnsiBold, kAnsiAccent), useAnsiColors);
    art.append(L"\r\n");
    AppendStyledText(art, LR"(   _/\     \__   )", kAnsiWarm, useAnsiColors);
    AppendStyledText(art, L"Fast local index and search daemon", kAnsiDim, useAnsiColors);
    art.append(L"\r\n");
    AppendStyledText(art, LR"(   \__  /\    _\ )", kAnsiWarm, useAnsiColors);
    art.append(L"\r\n");
    AppendStyledText(art, LR"(      \_/  \___/ )", kAnsiWarm, useAnsiColors);
    art.append(L"\r\n");
    return art;
}

[[nodiscard]] std::wstring GetForegroundRunModeText(uint32_t maxRequestsBeforeExit, uint32_t disconnectAfterBatches)
{
    if (maxRequestsBeforeExit == 0u && disconnectAfterBatches == 0u)
    {
        return L"Normal foreground run";
    }

    std::wstring mode = L"Test run";
    if (maxRequestsBeforeExit != 0u)
    {
        mode.append(std::format(L" | max requests={}", maxRequestsBeforeExit));
    }
    if (disconnectAfterBatches != 0u)
    {
        mode.append(std::format(L" | disconnect after {} batch(es)", disconnectAfterBatches));
    }
    return mode;
}

[[nodiscard]] std::wstring BuildForegroundStorePathText(const LocalSearchIndexCore::PersistentStoreInfo& storeInfo)
{
    if (! storeInfo.primaryPath.empty())
    {
        if (storeInfo.writeAheadLogPath.empty())
        {
            return std::format(L"{} ({})", storeInfo.primaryPath, FormatCompactBytes(storeInfo.primaryBytes));
        }

        return std::format(L"{} ({}) | WAL {} ({})",
                           storeInfo.primaryPath,
                           FormatCompactBytes(storeInfo.primaryBytes),
                           storeInfo.writeAheadLogPath,
                           FormatCompactBytes(storeInfo.writeAheadLogBytes));
    }

    if (storeInfo.kind == LocalSearchIndexCore::PersistentStoreKind::SnapshotBinary)
    {
        return L"Per-volume .bin snapshots";
    }

    return L"-";
}

[[nodiscard]] std::wstring BuildForegroundMaintenanceText(const LocalSearchIndexCore::PersistentStoreInfo& storeInfo)
{
    if (! storeInfo.autoCheckpointEnabled && ! storeInfo.autoCompactionEnabled)
    {
        return L"No automatic compaction";
    }

    std::wstring text;
    if (storeInfo.autoCheckpointEnabled)
    {
        text.append(std::format(L"checkpoint {}", FormatCompactBytes(storeInfo.autoCheckpointTargetBytes)));
    }
    if (storeInfo.autoCompactionEnabled)
    {
        if (! text.empty())
        {
            text.append(L" | ");
        }

        text.append(
            std::format(L"incremental compact {}% / {}", storeInfo.autoCompactionFragmentationPercent, FormatCompactBytes(storeInfo.autoCompactionMinBytes)));
    }
    return text;
}

[[nodiscard]] std::wstring FormatUtcHistoryValue(std::wstring_view value) noexcept
{
    return value.empty() ? L"never" : std::wstring(value);
}

[[nodiscard]] std::wstring BuildForegroundMaintenanceStateText(const LocalSearchIndexCore::PersistentStoreInfo& storeInfo,
                                                               const bool maintenanceQueued,
                                                               const bool maintenanceRunning)
{
    if (storeInfo.kind != LocalSearchIndexCore::PersistentStoreKind::Sqlite)
    {
        return L"Not applicable";
    }

    const std::wstring maintenanceState = maintenanceRunning ? L"running" : (maintenanceQueued ? L"queued" : L"idle");
    return std::format(L"state={} | pages={} free={} | checkpoint={} | compact={}",
                       maintenanceState,
                       storeInfo.pageCount,
                       storeInfo.freelistPageCount,
                       FormatUtcHistoryValue(storeInfo.lastCheckpointUtc),
                       FormatUtcHistoryValue(storeInfo.lastCompactionUtc));
}

[[nodiscard]] std::wstring BuildForegroundReadinessText(const LocalSearchIndexCore::PersistentStoreInfo& storeInfo)
{
    if (storeInfo.kind != LocalSearchIndexCore::PersistentStoreKind::Sqlite)
    {
        return L"Not applicable";
    }

    if (! storeInfo.inspectionSucceeded)
    {
        return L"Store not inspected yet";
    }

    return std::format(L"{} | volumes={} entries={} legacy-imports={}",
                       storeInfo.readyForQueryCutover ? L"Ready for query cutover" : L"Pending backfill/rebuild",
                       storeInfo.indexedVolumeCount,
                       storeInfo.indexedEntryCount,
                       storeInfo.legacyImportVolumeCount);
}

[[nodiscard]] std::wstring BuildForegroundStartupBanner(std::wstring_view pipeName,
                                                        std::wstring_view storageRoot,
                                                        const LocalSearchIndexCore::PersistentStoreInfo& storeInfo,
                                                        uint32_t protocolVersion,
                                                        uint32_t maxRequestsBeforeExit,
                                                        uint32_t disconnectAfterBatches,
                                                        bool renderDashboard,
                                                        bool useAnsiColors)
{
    std::wstring banner;
    AppendSeparator(banner, useAnsiColors);
    banner.append(BuildForegroundAsciiArt(useAnsiColors));
    AppendStyledText(
        banner, std::format(L"RedSalamander Search Service | {} | pid {}", GetBuildFlavorText(), ::GetCurrentProcessId()), kAnsiAccent, useAnsiColors);
    banner.append(L"\r\n");
    AppendSeparator(banner, useAnsiColors);
    AppendDashboardField(banner, L"Build", GetBuildFlavorText(), useAnsiColors, kAnsiInfo);
    AppendDashboardField(banner, L"Service", SearchServiceBroker::kServiceName, useAnsiColors, kAnsiAccent);
    AppendDashboardField(banner, L"Pipe", pipeName, useAnsiColors);
    AppendDashboardField(banner, L"Storage", storageRoot, useAnsiColors);
    AppendDashboardField(banner, L"Backend", LocalSearchIndexCore::GetPersistentStoreKindText(storeInfo.kind), useAnsiColors, kAnsiAccent);
    AppendDashboardField(banner, L"Store", BuildForegroundStorePathText(storeInfo), useAnsiColors);
    AppendDashboardField(banner, L"Maintenance", BuildForegroundMaintenanceText(storeInfo), useAnsiColors, kAnsiInfo);
    AppendDashboardField(banner, L"MaintState", BuildForegroundMaintenanceStateText(storeInfo, false, false), useAnsiColors, kAnsiDim);
    AppendDashboardField(
        banner, L"Readiness", BuildForegroundReadinessText(storeInfo), useAnsiColors, storeInfo.readyForQueryCutover ? kAnsiAccent : kAnsiWarm);
    AppendDashboardField(banner,
                         L"Warmup",
                         storeInfo.kind == LocalSearchIndexCore::PersistentStoreKind::Sqlite ? L"Awaiting startup state" : L"Disabled",
                         useAnsiColors,
                         storeInfo.kind == LocalSearchIndexCore::PersistentStoreKind::Sqlite ? kAnsiInfo : kAnsiDim);
    AppendDashboardField(banner, L"Protocol", std::format(L"{}", protocolVersion), useAnsiColors, kAnsiInfo);
    AppendDashboardField(banner, L"Mode", GetForegroundRunModeText(maxRequestsBeforeExit, disconnectAfterBatches), useAnsiColors, kAnsiInfo);
    AppendDashboardField(banner,
                         L"Output",
                         renderDashboard ? (useAnsiColors ? L"Interactive FTXUI dashboard (ANSI color)" : L"Interactive console dashboard")
                                         : L"Redirected lifecycle log",
                         useAnsiColors,
                         renderDashboard ? kAnsiAccent : kAnsiDim);
    AppendDashboardField(banner, L"Controls", L"Press Ctrl+C to stop.", useAnsiColors, kAnsiWarm);
    AppendSeparator(banner, useAnsiColors);
    return banner;
}

void InitializeForegroundConsole(ForegroundConsoleSession& session) noexcept
{
    const HANDLE stdoutHandle = ::GetStdHandle(STD_OUTPUT_HANDLE);
    if (IsConsoleHandle(stdoutHandle))
    {
        session.renderDashboard = true;
        static_cast<void>(::SetConsoleTitleW(L"RedSalamander Search Service - Foreground"));
        session.useAnsiColors     = TryEnableVirtualTerminal(stdoutHandle, session.originalOutputMode);
        session.restoreOutputMode = session.useAnsiColors;
        return;
    }

    if (HasUsableStdHandle(stdoutHandle) || HasUsableStdHandle(::GetStdHandle(STD_ERROR_HANDLE)))
    {
        session.renderDashboard = false;
        return;
    }

    bool hasConsole = ::GetConsoleWindow() != nullptr;
    if (! hasConsole)
    {
        if (::AttachConsole(ATTACH_PARENT_PROCESS) != FALSE)
        {
            hasConsole = true;
        }
        else
        {
            const DWORD error = ::GetLastError();
            if (error == ERROR_ACCESS_DENIED)
            {
                hasConsole = true;
            }
            else if (::AllocConsole() != FALSE)
            {
                hasConsole = true;
            }
        }
    }

    if (! hasConsole)
    {
        return;
    }

    session.stdinHandle.reset(::CreateFileW(L"CONIN$", GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0u, nullptr));
    session.stdoutHandle.reset(
        ::CreateFileW(L"CONOUT$", GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0u, nullptr));
    session.stderrHandle.reset(
        ::CreateFileW(L"CONOUT$", GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0u, nullptr));

    if (session.stdinHandle)
    {
        static_cast<void>(::SetStdHandle(STD_INPUT_HANDLE, session.stdinHandle.get()));
    }
    if (session.stdoutHandle)
    {
        static_cast<void>(::SetStdHandle(STD_OUTPUT_HANDLE, session.stdoutHandle.get()));
    }
    if (session.stderrHandle)
    {
        static_cast<void>(::SetStdHandle(STD_ERROR_HANDLE, session.stderrHandle.get()));
    }

    session.renderDashboard = IsConsoleHandle(session.GetOutputHandle());
    if (session.renderDashboard)
    {
        static_cast<void>(::SetConsoleTitleW(L"RedSalamander Search Service - Foreground"));
        session.useAnsiColors     = TryEnableVirtualTerminal(session.GetOutputHandle(), session.originalOutputMode);
        session.restoreOutputMode = session.useAnsiColors;
    }
}

[[nodiscard]] std::wstring BuildForegroundLogDetails(SearchServiceBroker::ServerEventType eventType,
                                                     const ForegroundRequestSnapshot& requestSnapshot,
                                                     const ForegroundWarmupSnapshot& warmupSnapshot,
                                                     const HRESULT hr) noexcept
{
    std::wstring details;

    if (eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_PROGRESS ||
        eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_FAILED ||
        eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_COMPLETED)
    {
        details.append(std::format(L" warmup={}/{} failed={}", warmupSnapshot.completedRoots, warmupSnapshot.totalRoots, warmupSnapshot.failedRoots));
        const std::wstring& rootText =
            (eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_FAILED && ! warmupSnapshot.lastFailureRoot.empty())
                ? warmupSnapshot.lastFailureRoot
                : warmupSnapshot.currentRoot;
        if (! rootText.empty())
        {
            details.append(std::format(L" root=\"{}\"", rootText));
        }
        if (warmupSnapshot.hasFailure)
        {
            details.append(std::format(L" lastFailure=0x{:08X}", static_cast<unsigned long>(warmupSnapshot.lastFailureHr)));
        }
        return details;
    }

    if (requestSnapshot.requestType == SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE)
    {
        if (eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_CLIENT_CONNECTED)
        {
            details.assign(L" awaiting-request-header");
        }
        else if ((eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_QUEUED ||
                  eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_RUNNING ||
                  eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_COMPLETED) &&
                 ! requestSnapshot.currentPath.empty())
        {
            details.assign(std::format(L" store=\"{}\"", requestSnapshot.currentPath));
        }
        return details;
    }

    details.append(std::format(L" req={}", GetRequestTypeText(requestSnapshot.requestType)));
    if (requestSnapshot.requestType == SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_QUERY)
    {
        details.append(std::format(L" mode=\"{}\"", FormatQueryModeSummary(requestSnapshot.requestNameMode, requestSnapshot.requestFlags)));
        if (! requestSnapshot.namePattern.empty())
        {
            details.append(std::format(L" pattern=\"{}\"", requestSnapshot.namePattern));
        }
        details.append(std::format(L" phase=\"{}\"", GetSearchPhaseText(requestSnapshot.requestPhase)));
    }

    if (! requestSnapshot.rootPath.empty())
    {
        details.append(std::format(L" root=\"{}\"", requestSnapshot.rootPath));
    }
    if (! requestSnapshot.currentPath.empty() && requestSnapshot.currentPath != requestSnapshot.rootPath)
    {
        details.append(std::format(L" path=\"{}\"", requestSnapshot.currentPath));
    }

    if (requestSnapshot.requestType == SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_QUERY)
    {
        details.append(std::format(L" scan=dirs:{} files:{} candidates:{} matches:{} batches:{}",
                                   requestSnapshot.scannedDirectories,
                                   requestSnapshot.scannedFiles,
                                   requestSnapshot.candidateFiles,
                                   requestSnapshot.matchedEntries,
                                   requestSnapshot.batchesSent));

        if (eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_QUERY_COMPLETED ||
            eventType == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_REQUEST_HANDLED || FAILED(hr))
        {
            details.append(std::format(L" result=store:{} memory:{} ready:{}ms query:{}ms",
                                       FormatCompactBytes(requestSnapshot.snapshotFileBytes),
                                       FormatCompactBytes(requestSnapshot.estimatedMemoryBytes),
                                       requestSnapshot.ensureReadyDurationMs,
                                       requestSnapshot.executeQueryDurationMs));
        }

        details.append(std::format(L" db=\"{}\" sync=\"{}\" search=\"{}\"",
                                   BuildForegroundStoreStateText(requestSnapshot),
                                   BuildForegroundSyncStateText(requestSnapshot),
                                   BuildForegroundExecutionModeText(requestSnapshot)));
    }

    return details;
}

void PrintForegroundLogLine(SearchServiceBroker::ServerEventType eventType,
                            uint32_t totalConnections,
                            uint32_t handledRequests,
                            HRESULT hr,
                            std::chrono::steady_clock::time_point startTime,
                            const ForegroundRequestSnapshot& requestSnapshot,
                            const ForegroundWarmupSnapshot& warmupSnapshot,
                            bool errorOutput) noexcept
{
    WriteConsoleText(std::format(L"[{}] {:<18} clients={} requests={} hr=0x{:08X}{}\r\n",
                                 FormatElapsedTime(std::chrono::steady_clock::now() - startTime),
                                 GetForegroundEventText(eventType, hr),
                                 totalConnections,
                                 handledRequests,
                                 static_cast<unsigned long>(hr),
                                 BuildForegroundLogDetails(eventType, requestSnapshot, warmupSnapshot, hr)),
                     errorOutput);
}

void STDMETHODCALLTYPE ForegroundServerEventCallback(const SearchServiceBroker::ServerEvent* event, void* cookie) noexcept
{
    if (event == nullptr || cookie == nullptr)
    {
        return;
    }

    auto& state = *static_cast<ForegroundConsoleState*>(cookie);
    state.lastEvent.store(static_cast<uint32_t>(event->type), std::memory_order_relaxed);
    state.lastHr.store(event->hr, std::memory_order_relaxed);
    state.handledRequests.store(event->handledRequests, std::memory_order_relaxed);
    state.seenEventMask.fetch_or(GetForegroundEventBit(event->type), std::memory_order_relaxed);
    state.uiRevision.fetch_add(1u, std::memory_order_relaxed);

    if (event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_CLIENT_CONNECTED)
    {
        state.totalConnections.fetch_add(1u, std::memory_order_relaxed);
    }
    if (event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_QUEUED)
    {
        state.maintenanceQueued.store(true, std::memory_order_relaxed);
        state.maintenanceRunning.store(false, std::memory_order_relaxed);
    }
    else if (event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_RUNNING)
    {
        state.maintenanceQueued.store(false, std::memory_order_relaxed);
        state.maintenanceRunning.store(true, std::memory_order_relaxed);
    }
    else if (event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_COMPLETED)
    {
        state.maintenanceQueued.store(false, std::memory_order_relaxed);
        state.maintenanceRunning.store(false, std::memory_order_relaxed);
    }

    const uint64_t nowTickMs = ::GetTickCount64();
    if (event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_CLIENT_CONNECTED)
    {
        state.activeRequestStartedTickMs.store(0u, std::memory_order_relaxed);
    }
    else if (event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_REQUEST_RECEIVED &&
             event->requestType != SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE)
    {
        state.activeRequestStartedTickMs.store(nowTickMs, std::memory_order_relaxed);
    }
    else if (event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_REQUEST_HANDLED ||
             event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STOPPED ||
             event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STOPPING)
    {
        state.activeRequestStartedTickMs.store(0u, std::memory_order_relaxed);
    }

    if (event->type != SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_WAITING_FOR_CLIENT &&
        event->type != SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING)
    {
        state.lastActivityEvent.store(static_cast<uint32_t>(event->type), std::memory_order_relaxed);
        state.lastActivityHr.store(event->hr, std::memory_order_relaxed);
        state.lastActivityRequests.store(event->handledRequests, std::memory_order_relaxed);
        state.lastActivityTickMs.store(nowTickMs, std::memory_order_relaxed);
    }

    {
        std::scoped_lock lock(state.detailsMutex);
        if (event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_CLIENT_CONNECTED)
        {
            state.requestType            = SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE;
            state.requestPhase           = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
            state.warningFlags           = FILESYSTEM_SEARCH_WARNING_NONE;
            state.requestFlags           = FILESYSTEM_SEARCH_NONE;
            state.requestNameMode        = FILESYSTEM_SEARCH_NAME_DISABLED;
            state.storeState             = static_cast<uint32_t>(LocalSearchIndexCore::StoreState::Unknown);
            state.syncPhase              = static_cast<uint32_t>(LocalSearchIndexCore::SyncPhase::Idle);
            state.queryExecutionMode     = static_cast<uint32_t>(LocalSearchIndexCore::QueryExecutionMode::Unknown);
            state.fallbackReason         = static_cast<uint32_t>(LocalSearchIndexCore::FallbackReason::None);
            state.completedRoots         = 0u;
            state.totalRoots             = 0u;
            state.batchesSent            = 0u;
            state.scannedDirectories     = 0u;
            state.scannedFiles           = 0u;
            state.candidateFiles         = 0u;
            state.matchedEntries         = 0u;
            state.snapshotFileBytes      = 0u;
            state.estimatedMemoryBytes   = 0u;
            state.ensureReadyDurationMs  = 0u;
            state.executeQueryDurationMs = 0u;
            CopyFixedText(state.rootPath, nullptr);
            CopyFixedText(state.namePattern, nullptr);
            CopyFixedText(state.activeRoot, nullptr);
            CopyFixedText(state.currentPath, nullptr);
        }

        if (event->requestType != SearchServiceBroker::SEARCH_SERVICE_SERVER_REQUEST_NONE || event->rootPath != nullptr || event->currentPath != nullptr)
        {
            state.requestType            = event->requestType;
            state.requestPhase           = event->phase;
            state.warningFlags           = event->warningFlags;
            state.requestFlags           = event->requestFlags;
            state.requestNameMode        = event->requestNameMode;
            state.storeState             = event->storeState;
            state.syncPhase              = event->syncPhase;
            state.queryExecutionMode     = event->queryExecutionMode;
            state.fallbackReason         = event->fallbackReason;
            state.completedRoots         = event->completedRoots;
            state.totalRoots             = event->totalRoots;
            state.batchesSent            = event->batchesSent;
            state.scannedDirectories     = event->scannedDirectories;
            state.scannedFiles           = event->scannedFiles;
            state.candidateFiles         = event->candidateFiles;
            state.matchedEntries         = event->matchedEntries;
            state.snapshotFileBytes      = event->snapshotFileBytes;
            state.estimatedMemoryBytes   = event->estimatedMemoryBytes;
            state.ensureReadyDurationMs  = event->ensureReadyDurationMs;
            state.executeQueryDurationMs = event->executeQueryDurationMs;
            CopyFixedText(state.rootPath, event->rootPath);
            CopyFixedText(state.namePattern, event->namePattern);
            CopyFixedText(state.activeRoot, event->activeRoot);
            CopyFixedText(state.currentPath, event->currentPath);
        }

        if (event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_PROGRESS ||
            event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_FAILED ||
            event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_COMPLETED)
        {
            state.startupWarmupEnabled        = true;
            state.startupWarmupRunning        = event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_PROGRESS;
            state.startupWarmupTotalRoots     = event->startupWarmupTotalRoots;
            state.startupWarmupCompletedRoots = event->startupWarmupCompletedRoots;
            state.startupWarmupFailedRoots    = event->startupWarmupFailedRoots;
            CopyFixedText(state.startupWarmupCurrentRoot, event->currentPath);
            if (event->type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_FAILED)
            {
                state.startupWarmupHasFailure    = true;
                state.startupWarmupLastFailureHr = event->hr;
                CopyFixedText(state.startupWarmupLastFailureRoot,
                              event->startupWarmupLastFailureRoot != nullptr ? event->startupWarmupLastFailureRoot : event->currentPath);
            }
            else if (FAILED(event->startupWarmupLastFailureHr) || event->startupWarmupLastFailureRoot != nullptr)
            {
                state.startupWarmupHasFailure    = FAILED(event->startupWarmupLastFailureHr) || state.startupWarmupFailedRoots != 0u;
                state.startupWarmupLastFailureHr = event->startupWarmupLastFailureHr;
                CopyFixedText(state.startupWarmupLastFailureRoot, event->startupWarmupLastFailureRoot);
            }
        }

        if (event->type != SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_STARTING)
        {
            ForegroundEventRecord record{};
            record.type                                   = event->type;
            record.hr                                     = event->hr;
            record.tickMs                                 = nowTickMs;
            record.totalConnections                       = state.totalConnections.load(std::memory_order_relaxed);
            record.handledRequests                        = event->handledRequests;
            record.requestSnapshot.requestType            = state.requestType;
            record.requestSnapshot.requestPhase           = state.requestPhase;
            record.requestSnapshot.warningFlags           = state.warningFlags;
            record.requestSnapshot.requestFlags           = state.requestFlags;
            record.requestSnapshot.requestNameMode        = state.requestNameMode;
            record.requestSnapshot.storeState             = static_cast<LocalSearchIndexCore::StoreState>(state.storeState);
            record.requestSnapshot.syncPhase              = static_cast<LocalSearchIndexCore::SyncPhase>(state.syncPhase);
            record.requestSnapshot.queryExecutionMode     = static_cast<LocalSearchIndexCore::QueryExecutionMode>(state.queryExecutionMode);
            record.requestSnapshot.fallbackReason         = static_cast<LocalSearchIndexCore::FallbackReason>(state.fallbackReason);
            record.requestSnapshot.completedRoots         = state.completedRoots;
            record.requestSnapshot.totalRoots             = state.totalRoots;
            record.requestSnapshot.batchesSent            = state.batchesSent;
            record.requestSnapshot.scannedDirectories     = state.scannedDirectories;
            record.requestSnapshot.scannedFiles           = state.scannedFiles;
            record.requestSnapshot.candidateFiles         = state.candidateFiles;
            record.requestSnapshot.matchedEntries         = state.matchedEntries;
            record.requestSnapshot.snapshotFileBytes      = state.snapshotFileBytes;
            record.requestSnapshot.estimatedMemoryBytes   = state.estimatedMemoryBytes;
            record.requestSnapshot.ensureReadyDurationMs  = state.ensureReadyDurationMs;
            record.requestSnapshot.executeQueryDurationMs = state.executeQueryDurationMs;
            record.requestSnapshot.rootPath.assign(ViewFixedText(state.rootPath));
            record.requestSnapshot.namePattern.assign(ViewFixedText(state.namePattern));
            record.requestSnapshot.activeRoot.assign(ViewFixedText(state.activeRoot));
            record.requestSnapshot.currentPath.assign(ViewFixedText(state.currentPath));
            record.warmupSnapshot.enabled        = state.startupWarmupEnabled;
            record.warmupSnapshot.running        = state.startupWarmupRunning;
            record.warmupSnapshot.totalRoots     = state.startupWarmupTotalRoots;
            record.warmupSnapshot.completedRoots = state.startupWarmupCompletedRoots;
            record.warmupSnapshot.failedRoots    = state.startupWarmupFailedRoots;
            record.warmupSnapshot.hasFailure     = state.startupWarmupHasFailure;
            record.warmupSnapshot.lastFailureHr  = state.startupWarmupLastFailureHr;
            record.warmupSnapshot.currentRoot.assign(ViewFixedText(state.startupWarmupCurrentRoot));
            record.warmupSnapshot.lastFailureRoot.assign(ViewFixedText(state.startupWarmupLastFailureRoot));
            state.pendingEvents.push_back(std::move(record));
            constexpr size_t kMaxPendingForegroundEvents = 512u;
            while (state.pendingEvents.size() > kMaxPendingForegroundEvents)
            {
                state.pendingEvents.pop_front();
            }
        }
    }

    const HANDLE wakeEvent = g_foregroundUiWakeEvent.load(std::memory_order_acquire);
    if (wakeEvent != nullptr)
    {
        static_cast<void>(::SetEvent(wakeEvent));
    }
}

BOOL WINAPI ForegroundConsoleCtrlHandler(DWORD controlType) noexcept
{
    switch (controlType)
    {
        case CTRL_C_EVENT:
        case CTRL_BREAK_EVENT:
        case CTRL_CLOSE_EVENT:
        case CTRL_SHUTDOWN_EVENT:
            if (ForegroundConsoleState* const consoleState = g_foregroundConsoleState.load(std::memory_order_acquire); consoleState != nullptr)
            {
                consoleState->stopRequested.store(true, std::memory_order_relaxed);
            }
            if (const HANDLE stopEvent = g_foregroundStopEvent.load(std::memory_order_acquire); stopEvent != nullptr)
            {
                static_cast<void>(::SetEvent(stopEvent));
            }
            if (const HANDLE wakeEvent = g_foregroundUiWakeEvent.load(std::memory_order_acquire); wakeEvent != nullptr)
            {
                static_cast<void>(::SetEvent(wakeEvent));
            }
            return TRUE;

        default: return FALSE;
    }
}

void PrintHelpText() noexcept
{
    const std::wstring helpText =
        std::format(L"RedSalamanderSearchService\r\n"
                    L"\r\n"
                    L"Usage:\r\n"
                    L"  {0} [options]\r\n"
                    L"\r\n"
                    L"Options:\r\n"
                    L"  -h, --help, /?                 Show this help.\r\n"
                    L"  --run-foreground               Run in the current terminal instead of the SCM.\r\n"
                    L"  --compact                      Run offline SQLite checkpoint/vacuum maintenance and exit.\r\n"
                    L"  --request-compact              Ask the running service to compact its live SQLite store.\r\n"
                    L"  --register                     Register this executable as the Windows service '{1}'.\r\n"
                    L"  --unregister                   Unregister the Windows service '{1}'.\r\n"
                    L"  --pipe-name=NAME               Override the pipe name for foreground mode or --request-compact.\r\n"
                    L"  --storage-root=PATH            Override the persistent-store root for foreground mode or --compact.\r\n"
                    L"  --store-backend=KIND           Select the persistent store backend for foreground mode or --compact: snapshot or sqlite.\r\n"
                    L"  --sqlite-path=PATH             Override the SQLite database path for foreground mode or --compact.\r\n"
                    L"  --protocol-version=N           Override the protocol version for foreground mode.\r\n"
                    L"  --max-requests=N               Exit after N handled requests in foreground mode (0 = unlimited).\r\n"
                    L"  --disconnect-after-batches=N   Disconnect after N candidate batches in foreground mode (0 = never).\r\n"
                    L"\r\n"
                    L"Defaults for this build:\r\n"
                    L"  Service name: {1}\r\n"
                    L"  Pipe name:    {2}\r\n"
                    L"  Store:        sqlite-v2\r\n"
                    L"\r\n"
                    L"Notes:\r\n"
                    L"  --register and --unregister typically require an elevated terminal.\r\n"
                    L"  Only one active '{1}' instance may run at a time for this build.\r\n"
                    L"  --run-foreground prints a startup banner with PID/build/mode details, shows a full-screen interactive dashboard\r\n"
                    L"                   with live status and history pages on a VT-capable console, and emits readable log lines when\r\n"
                    L"                   output is redirected.\r\n"
                    L"  --compact targets the SQLite store, acquires the single-instance guard, truncates WAL, runs VACUUM,\r\n"
                    L"            and records the last checkpoint/compaction timestamps.\r\n"
                    L"  --request-compact talks to the running service over the named pipe, runs live maintenance in-process,\r\n"
                    L"                    and prints the refreshed store state after the request completes.\r\n"
                    L"  ANSI colors are enabled automatically when the active terminal supports virtual terminal sequences.\r\n"
                    L"  Use 1/2 to switch pages, F to follow the latest event, arrows to browse history, and Ctrl+C to stop.\r\n"
                    L"\r\n"
                    L"Examples:\r\n"
                    L"  {0} --help\r\n"
                    L"  {0} --run-foreground\r\n"
                    L"  {0} --run-foreground --storage-root=C:\\Temp\\RedSalamander\\SearchIndex\r\n"
                    L"  {0} --run-foreground --store-backend=snapshot\r\n"
                    L"  {0} --compact\r\n"
                    L"  {0} --compact --sqlite-path=C:\\Temp\\RedSalamander\\SearchIndex\\index-v2.sqlite3\r\n"
                    L"  {0} --request-compact\r\n"
                    L"  {0} --register\r\n"
                    L"  {0} --unregister\r\n",
                    GetExecutableLeafName(),
                    SearchServiceBroker::kServiceName,
                    SearchServiceBroker::GetDefaultPipeName());
    WriteConsoleText(helpText, false);
}

void PrintUsageError(std::wstring_view message) noexcept
{
    WriteConsoleText(std::format(L"{}\r\nUse --help for usage.\r\n", message), true);
}

HRESULT TryGetCurrentExecutablePath(std::wstring& outPath) noexcept
{
    outPath.assign(260u, L'\0');
    for (;;)
    {
        const DWORD written = ::GetModuleFileNameW(nullptr, outPath.data(), static_cast<DWORD>(outPath.size()));
        if (written == 0u)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        if (written < outPath.size())
        {
            outPath.resize(written);
            return S_OK;
        }

        if (outPath.size() >= 32768u)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }

        outPath.resize(outPath.size() * 2u);
    }
}

HRESULT QueryCurrentServiceStatus(SC_HANDLE service, SERVICE_STATUS_PROCESS& outStatus) noexcept
{
    outStatus         = {};
    DWORD bytesNeeded = 0u;
    if (::QueryServiceStatusEx(service, SC_STATUS_PROCESS_INFO, reinterpret_cast<BYTE*>(&outStatus), static_cast<DWORD>(sizeof(outStatus)), &bytesNeeded) == 0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    return S_OK;
}

HRESULT StopServiceAndWaitForDeletion(SC_HANDLE service) noexcept
{
    SERVICE_STATUS_PROCESS status{};
    HRESULT hr = QueryCurrentServiceStatus(service, status);
    if (FAILED(hr))
    {
        return hr;
    }

    if (status.dwCurrentState == SERVICE_STOPPED)
    {
        return S_OK;
    }

    if (status.dwCurrentState != SERVICE_STOP_PENDING)
    {
        SERVICE_STATUS serviceStatus{};
        if (::ControlService(service, SERVICE_CONTROL_STOP, &serviceStatus) == 0)
        {
            const DWORD error = ::GetLastError();
            if (error != ERROR_SERVICE_NOT_ACTIVE)
            {
                return HRESULT_FROM_WIN32(error);
            }
        }
    }

    const ULONGLONG deadline = ::GetTickCount64() + 15000u;
    while (status.dwCurrentState != SERVICE_STOPPED)
    {
        if (::GetTickCount64() >= deadline)
        {
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }

        DWORD waitMs = status.dwWaitHint / 10u;
        waitMs       = std::max<DWORD>(100u, std::min<DWORD>(waitMs, 1000u));
        ::Sleep(waitMs);

        hr = QueryCurrentServiceStatus(service, status);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    return S_OK;
}

CommandResult RegisterCurrentService() noexcept
{
    CommandResult result{};

    try
    {
        std::wstring executablePath;
        HRESULT hr = TryGetCurrentExecutablePath(executablePath);
        if (FAILED(hr))
        {
            result.hr      = hr;
            result.message = std::format(L"Failed to resolve the current executable path. hr=0x{:08X}", static_cast<unsigned long>(hr));
            return result;
        }

        wil::unique_schandle scm(::OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT | SC_MANAGER_CREATE_SERVICE));
        if (! scm)
        {
            const DWORD error = ::GetLastError();
            result.hr         = HRESULT_FROM_WIN32(error);
            result.message    = std::format(L"OpenSCManagerW failed. error={}", error);
            return result;
        }

        const std::wstring binaryPath  = std::format(L"\"{}\"", executablePath);
        const std::wstring displayName = GetServiceDisplayName();

        wil::unique_schandle service(::CreateServiceW(scm.get(),
                                                      SearchServiceBroker::kServiceName,
                                                      displayName.c_str(),
                                                      SERVICE_CHANGE_CONFIG | SERVICE_QUERY_STATUS | DELETE,
                                                      SERVICE_WIN32_OWN_PROCESS,
                                                      SERVICE_AUTO_START,
                                                      SERVICE_ERROR_NORMAL,
                                                      binaryPath.c_str(),
                                                      nullptr,
                                                      nullptr,
                                                      nullptr,
                                                      nullptr,
                                                      nullptr));

        bool updatedExisting = false;
        if (! service)
        {
            const DWORD error = ::GetLastError();
            if (error != ERROR_SERVICE_EXISTS)
            {
                result.hr      = HRESULT_FROM_WIN32(error);
                result.message = std::format(L"CreateServiceW failed for '{}'. error={}", SearchServiceBroker::kServiceName, error);
                return result;
            }

            service.reset(::OpenServiceW(scm.get(), SearchServiceBroker::kServiceName, SERVICE_CHANGE_CONFIG | SERVICE_QUERY_STATUS | DELETE));
            if (! service)
            {
                const DWORD openError = ::GetLastError();
                result.hr             = HRESULT_FROM_WIN32(openError);
                result.message        = std::format(L"OpenServiceW failed for '{}'. error={}", SearchServiceBroker::kServiceName, openError);
                return result;
            }

            if (::ChangeServiceConfigW(service.get(),
                                       SERVICE_NO_CHANGE,
                                       SERVICE_AUTO_START,
                                       SERVICE_ERROR_NORMAL,
                                       binaryPath.c_str(),
                                       nullptr,
                                       nullptr,
                                       nullptr,
                                       nullptr,
                                       nullptr,
                                       displayName.c_str()) == 0)
            {
                const DWORD configError = ::GetLastError();
                result.hr               = HRESULT_FROM_WIN32(configError);
                result.message          = std::format(L"ChangeServiceConfigW failed for '{}'. error={}", SearchServiceBroker::kServiceName, configError);
                return result;
            }

            updatedExisting = true;
        }

        std::wstring warningSuffix;
        std::wstring description = GetServiceDescription();
        SERVICE_DESCRIPTIONW serviceDescription{};
        serviceDescription.lpDescription = description.data();
        if (::ChangeServiceConfig2W(service.get(), SERVICE_CONFIG_DESCRIPTION, &serviceDescription) == 0)
        {
            const DWORD error = ::GetLastError();
            warningSuffix     = std::format(L" Warning: failed to update the service description (error={}).", error);
        }

        result.hr      = updatedExisting ? S_FALSE : S_OK;
        result.message = std::format(L"{} service '{}' using {}. Start it with: sc start {}.{}",
                                     updatedExisting ? L"Updated" : L"Registered",
                                     SearchServiceBroker::kServiceName,
                                     binaryPath,
                                     SearchServiceBroker::kServiceName,
                                     warningSuffix);
        return result;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"RedSalamanderSearchService: RegisterCurrentService failed with an unexpected std::exception.");
        result.hr      = E_FAIL;
        result.message = L"RegisterCurrentService failed with an unexpected std::exception.";
        return result;
    }
}

CommandResult UnregisterCurrentService() noexcept
{
    CommandResult result{};

    try
    {
        wil::unique_schandle scm(::OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT));
        if (! scm)
        {
            const DWORD error = ::GetLastError();
            result.hr         = HRESULT_FROM_WIN32(error);
            result.message    = std::format(L"OpenSCManagerW failed. error={}", error);
            return result;
        }

        wil::unique_schandle service(::OpenServiceW(scm.get(), SearchServiceBroker::kServiceName, SERVICE_STOP | SERVICE_QUERY_STATUS | DELETE));
        if (! service)
        {
            const DWORD error = ::GetLastError();
            if (error == ERROR_SERVICE_DOES_NOT_EXIST)
            {
                result.hr      = S_FALSE;
                result.message = std::format(L"Service '{}' is not registered.", SearchServiceBroker::kServiceName);
                return result;
            }

            result.hr      = HRESULT_FROM_WIN32(error);
            result.message = std::format(L"OpenServiceW failed for '{}'. error={}", SearchServiceBroker::kServiceName, error);
            return result;
        }

        HRESULT hr = StopServiceAndWaitForDeletion(service.get());
        if (FAILED(hr))
        {
            result.hr = hr;
            result.message =
                std::format(L"Failed to stop service '{}' before deletion. hr=0x{:08X}", SearchServiceBroker::kServiceName, static_cast<unsigned long>(hr));
            return result;
        }

        if (::DeleteService(service.get()) == 0)
        {
            const DWORD error = ::GetLastError();
            if (error != ERROR_SERVICE_MARKED_FOR_DELETE)
            {
                result.hr      = HRESULT_FROM_WIN32(error);
                result.message = std::format(L"DeleteService failed for '{}'. error={}", SearchServiceBroker::kServiceName, error);
                return result;
            }
        }

        result.hr      = S_OK;
        result.message = std::format(L"Unregistered service '{}'.", SearchServiceBroker::kServiceName);
        return result;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"RedSalamanderSearchService: UnregisterCurrentService failed with an unexpected std::exception.");
        result.hr      = E_FAIL;
        result.message = L"UnregisterCurrentService failed with an unexpected std::exception.";
        return result;
    }
}

CommandResult RunCompactStore(const ParsedArguments& parsed) noexcept
{
    CommandResult result{};

    if (parsed.persistentStoreKind != LocalSearchIndexCore::PersistentStoreKind::Sqlite)
    {
        result.hr      = E_INVALIDARG;
        result.message = L"--compact only supports the sqlite backend. Use --store-backend=sqlite or --sqlite-path=PATH.";
        return result;
    }

    wil::unique_handle instanceGuard;
    bool alreadyRunning      = false;
    const HRESULT instanceHr = AcquireSingleInstanceGuard(instanceGuard, alreadyRunning);
    if (FAILED(instanceHr))
    {
        result.hr      = instanceHr;
        result.message = std::format(L"Failed to acquire the single-instance guard before compaction. hr=0x{:08X}", static_cast<unsigned long>(instanceHr));
        return result;
    }
    if (alreadyRunning)
    {
        result.hr      = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        result.message = std::format(L"{} Stop it before running --compact.", BuildAlreadyRunningMessage());
        return result;
    }

    const LocalSearchIndexCore::PersistentStoreInfo storeInfo = LocalSearchIndexCore::GetPersistentStoreInfo(BuildRepositoryOptions(parsed));
    if (storeInfo.primaryPath.empty())
    {
        result.hr      = E_FAIL;
        result.message = L"Failed to resolve the SQLite database path for --compact.";
        return result;
    }

    SqliteIndexStore::ManualMaintenanceResult maintenance{};
    const HRESULT hr = SqliteIndexStore::RunManualMaintenance(storeInfo.primaryPath, &maintenance);
    if (FAILED(hr))
    {
        result.hr      = hr;
        result.message = std::format(L"Compaction failed for '{}'. hr=0x{:08X}", storeInfo.primaryPath, static_cast<unsigned long>(hr));
        return result;
    }

    const std::wstring checkpointUtc = maintenance.after.lastCheckpointUtc.empty() ? L"never" : maintenance.after.lastCheckpointUtc;
    const std::wstring compactionUtc = maintenance.after.lastCompactionUtc.empty() ? L"never" : maintenance.after.lastCompactionUtc;

    result.hr      = S_OK;
    result.message = std::format(L"Compaction completed for '{0}'. db {1} -> {2}, WAL {3} -> {4}, free pages {5} -> {6}, checkpoint={7}, compact={8}.",
                                 maintenance.after.databasePath,
                                 FormatCompactBytes(maintenance.before.databaseBytes),
                                 FormatCompactBytes(maintenance.after.databaseBytes),
                                 FormatCompactBytes(maintenance.before.writeAheadLogBytes),
                                 FormatCompactBytes(maintenance.after.writeAheadLogBytes),
                                 maintenance.before.freelistPageCount,
                                 maintenance.after.freelistPageCount,
                                 checkpointUtc,
                                 compactionUtc);
    return result;
}

CommandResult RequestServiceCompaction(const ParsedArguments& parsed) noexcept
{
    CommandResult result{};

    const std::wstring previousPipeOverride = GetEnvironmentVariableText(SearchServiceBroker::kPipeNameEnvVar);
    if (! parsed.pipeName.empty())
    {
        if (::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, parsed.pipeName.c_str()) == 0)
        {
            const DWORD error = ::GetLastError();
            result.hr         = HRESULT_FROM_WIN32(error);
            result.message    = std::format(L"Failed to apply --pipe-name override '{}'. error={}", parsed.pipeName, error);
            return result;
        }
    }
    const auto restorePipeOverride = wil::scope_exit([&]() noexcept
    {
        if (! parsed.pipeName.empty())
        {
            static_cast<void>(
                ::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str()));
        }
    });

    const HRESULT hr = SearchServiceBroker::RequestCompact();
    if (FAILED(hr))
    {
        result.hr      = hr;
        result.message = std::format(L"Live compaction request failed. hr=0x{:08X}. Start the service with the sqlite backend or use --compact offline.",
                                     static_cast<unsigned long>(hr));
        return result;
    }

    SearchServiceBroker::ServiceStatus status{};
    const HRESULT statusHr = SearchServiceBroker::GetStatus(status);
    if (FAILED(statusHr))
    {
        result.hr = S_OK;
        result.message =
            std::format(L"Live compaction request completed, but the follow-up status probe failed. hr=0x{:08X}", static_cast<unsigned long>(statusHr));
        return result;
    }

    result.hr      = S_OK;
    result.message = std::format(L"Live compaction completed for '{0}'. db={1}, WAL={2}, free pages={3}, checkpoint={4}, compact={5}.",
                                 status.persistentStorePath,
                                 FormatCompactBytes(status.persistentStoreBytes),
                                 FormatCompactBytes(status.writeAheadLogBytes),
                                 status.persistentStoreFreelistPageCount,
                                 status.lastCheckpointUtc.empty() ? L"never" : status.lastCheckpointUtc,
                                 status.lastCompactionUtc.empty() ? L"never" : status.lastCompactionUtc);
    return result;
}

CommandResult RunForegroundService(const ParsedArguments& parsed) noexcept
{
    CommandResult result{};

    wil::unique_handle instanceGuard;
    bool alreadyRunning      = false;
    const HRESULT instanceHr = AcquireSingleInstanceGuard(instanceGuard, alreadyRunning);
    if (FAILED(instanceHr))
    {
        result.hr      = instanceHr;
        result.message = std::format(
            L"Failed to create the single-instance guard for '{}'. hr=0x{:08X}", SearchServiceBroker::kServiceName, static_cast<unsigned long>(instanceHr));
        return result;
    }
    if (alreadyRunning)
    {
        result.hr      = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        result.message = BuildAlreadyRunningMessage();
        return result;
    }

    SearchServiceBroker::ServerOptions options{};
    options.pipeName               = parsed.pipeName;
    options.storageRootDirectory   = parsed.storageRootDirectory;
    options.persistentStoreKind    = parsed.persistentStoreKind;
    options.sqliteDatabasePath     = parsed.sqliteDatabasePath;
    options.protocolVersion        = parsed.protocolVersion;
    options.maxRequestsBeforeExit  = parsed.maxRequestsBeforeExit;
    options.disconnectAfterBatches = parsed.disconnectAfterBatches;

    const std::wstring pipeName    = parsed.pipeName.empty() ? SearchServiceBroker::GetDefaultPipeName() : parsed.pipeName;
    const std::wstring storageRoot = parsed.storageRootDirectory.empty() ? SearchServiceBroker::GetProgramDataSearchIndexRoot() : parsed.storageRootDirectory;
    const auto getStoreInfo        = [&]() noexcept
    {
        return LocalSearchIndexCore::GetPersistentStoreInfo(
            {.snapshotRootDirectory = storageRoot,
             .persistentStoreKind   = options.persistentStoreKind,
             .sqliteDatabasePath    = options.sqliteDatabasePath,
             .sqliteAuthoritative   = options.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite});
    };

    wil::unique_event stopEvent;
    stopEvent.create();
    if (! stopEvent)
    {
        result.hr      = E_OUTOFMEMORY;
        result.message = L"Failed to create the foreground stop event.";
        return result;
    }

    ForegroundConsoleState consoleState{};
    options.eventCallback = &ForegroundServerEventCallback;
    options.eventCookie   = &consoleState;

    SearchServiceBroker::ServerRunResult runResult{};
    std::atomic<long> serverHr{S_OK};
    std::atomic<bool> serverExited{false};

    ForegroundConsoleSession consoleSession{};
    InitializeForegroundConsole(consoleSession);
    const bool renderDashboard = consoleSession.renderDashboard && consoleSession.useAnsiColors;
    const bool useAnsiColors   = consoleSession.useAnsiColors;
    const auto startTime       = std::chrono::steady_clock::now();

    g_foregroundStopEvent.store(stopEvent.get(), std::memory_order_release);
    g_foregroundConsoleState.store(&consoleState, std::memory_order_release);
    static_cast<void>(::SetConsoleCtrlHandler(&ForegroundConsoleCtrlHandler, TRUE));
    const auto restoreConsoleHandler = wil::scope_exit([&] noexcept
    {
        static_cast<void>(::SetConsoleCtrlHandler(&ForegroundConsoleCtrlHandler, FALSE));
        g_foregroundConsoleState.store(nullptr, std::memory_order_release);
        g_foregroundStopEvent.store(nullptr, std::memory_order_release);
    });
    const auto restoreConsoleMode    = wil::scope_exit([&] noexcept
    {
        if (consoleSession.restoreOutputMode)
        {
            static_cast<void>(::SetConsoleMode(consoleSession.GetOutputHandle(), consoleSession.originalOutputMode));
        }
    });

    if (! renderDashboard)
    {
        WriteConsoleText(BuildForegroundStartupBanner(pipeName,
                                                      storageRoot,
                                                      getStoreInfo(),
                                                      parsed.protocolVersion,
                                                      parsed.maxRequestsBeforeExit,
                                                      parsed.disconnectAfterBatches,
                                                      renderDashboard,
                                                      useAnsiColors),
                         false);
    }

    std::jthread serverThread([&] noexcept
    {
        SearchServiceBroker::ServerRunResult localResult{};
        const HRESULT hr = SearchServiceBroker::RunServer(options, stopEvent.get(), &localResult);
        runResult        = localResult;
        serverHr.store(hr, std::memory_order_relaxed);
        serverExited.store(true, std::memory_order_release);
        if (const HANDLE wakeEvent = g_foregroundUiWakeEvent.load(std::memory_order_acquire); wakeEvent != nullptr)
        {
            static_cast<void>(::SetEvent(wakeEvent));
        }
    });

    if (renderDashboard)
    {
        ForegroundDashboardController dashboard(consoleState,
                                                pipeName,
                                                storageRoot,
                                                parsed.persistentStoreKind,
                                                parsed.sqliteDatabasePath,
                                                parsed.protocolVersion,
                                                parsed.maxRequestsBeforeExit,
                                                parsed.disconnectAfterBatches,
                                                startTime);
        dashboard.Run(serverExited);
    }
    else
    {
        while (! serverExited.load(std::memory_order_acquire))
        {
            std::deque<ForegroundEventRecord> pendingEvents;
            {
                std::scoped_lock lock(consoleState.detailsMutex);
                pendingEvents.swap(consoleState.pendingEvents);
            }

            for (const ForegroundEventRecord& pendingEvent : pendingEvents)
            {
                PrintForegroundLogLine(pendingEvent.type,
                                       pendingEvent.totalConnections,
                                       pendingEvent.handledRequests,
                                       pendingEvent.hr,
                                       startTime,
                                       pendingEvent.requestSnapshot,
                                       pendingEvent.warmupSnapshot,
                                       FAILED(pendingEvent.hr) || pendingEvent.type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_ERROR);
            }

            std::this_thread::sleep_for(std::chrono::milliseconds(200));
        }
    }

    const HRESULT hr = static_cast<HRESULT>(serverHr.load(std::memory_order_relaxed));
    if (! renderDashboard)
    {
        std::deque<ForegroundEventRecord> pendingEvents;
        {
            std::scoped_lock lock(consoleState.detailsMutex);
            pendingEvents.swap(consoleState.pendingEvents);
        }

        for (const ForegroundEventRecord& pendingEvent : pendingEvents)
        {
            PrintForegroundLogLine(pendingEvent.type,
                                   pendingEvent.totalConnections,
                                   pendingEvent.handledRequests,
                                   pendingEvent.hr,
                                   startTime,
                                   pendingEvent.requestSnapshot,
                                   pendingEvent.warmupSnapshot,
                                   FAILED(pendingEvent.hr) || pendingEvent.type == SearchServiceBroker::SEARCH_SERVICE_SERVER_EVENT_ERROR);
        }
    }

    result.hr = hr;
    result.message =
        SUCCEEDED(hr)
            ? std::format(L"Foreground service stopped after {} handled request(s).", runResult.handledRequests)
            : std::format(L"Foreground service failed after {} handled request(s): 0x{:08X}", runResult.handledRequests, static_cast<unsigned long>(hr));
    return result;
}

ParsedArguments ParseArguments() noexcept
{
    ParsedArguments parsed{};

    int argc = 0;
    wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
    if (! argv)
    {
        return parsed;
    }

    for (int index = 1; index < argc; ++index)
    {
        const std::wstring_view arg(argv.get()[index] ? argv.get()[index] : L"");
        if (arg == L"--help" || arg == L"-h" || arg == L"/?")
        {
            static_cast<void>(TrySetAction(parsed, StartupAction::ShowHelp, arg));
            continue;
        }

        if (arg == L"--run-foreground")
        {
            if (! TrySetAction(parsed, StartupAction::RunForeground, arg))
            {
                return parsed;
            }
            continue;
        }

        if (arg == L"--compact")
        {
            if (! TrySetAction(parsed, StartupAction::CompactStore, arg))
            {
                return parsed;
            }
            continue;
        }

        if (arg == L"--request-compact")
        {
            if (! TrySetAction(parsed, StartupAction::RequestCompactService, arg))
            {
                return parsed;
            }
            continue;
        }

        if (arg == L"--register")
        {
            if (! TrySetAction(parsed, StartupAction::RegisterService, arg))
            {
                return parsed;
            }
            continue;
        }

        if (arg == L"--unregister")
        {
            if (! TrySetAction(parsed, StartupAction::UnregisterService, arg))
            {
                return parsed;
            }
            continue;
        }

        if (arg.rfind(L"--pipe-name=", 0) == 0)
        {
            parsed.pipeName = std::wstring(arg.substr(std::wstring_view(L"--pipe-name=").size()));
            continue;
        }

        if (arg.rfind(L"--storage-root=", 0) == 0)
        {
            parsed.storageRootDirectory = std::wstring(arg.substr(std::wstring_view(L"--storage-root=").size()));
            continue;
        }

        if (arg.rfind(L"--store-backend=", 0) == 0)
        {
            const std::wstring_view value = arg.substr(std::wstring_view(L"--store-backend=").size());
            if (! TryParseStoreBackend(value, parsed.persistentStoreKind))
            {
                parsed.errorMessage = std::format(L"Invalid value for --store-backend: '{}'. Use snapshot or sqlite.", value);
                return parsed;
            }
            parsed.storeBackendExplicit = true;
            continue;
        }

        if (arg.rfind(L"--sqlite-path=", 0) == 0)
        {
            parsed.sqliteDatabasePath  = std::wstring(arg.substr(std::wstring_view(L"--sqlite-path=").size()));
            parsed.persistentStoreKind = LocalSearchIndexCore::PersistentStoreKind::Sqlite;
            continue;
        }

        if (arg.rfind(L"--protocol-version=", 0) == 0)
        {
            uint32_t value = 0u;
            if (! TryParseUnsigned(arg.substr(std::wstring_view(L"--protocol-version=").size()), value))
            {
                parsed.errorMessage = std::format(L"Invalid value for --protocol-version: '{}'.", arg);
                return parsed;
            }

            parsed.protocolVersion = value;
            continue;
        }

        if (arg.rfind(L"--max-requests=", 0) == 0)
        {
            uint32_t value = 0u;
            if (! TryParseUnsigned(arg.substr(std::wstring_view(L"--max-requests=").size()), value))
            {
                parsed.errorMessage = std::format(L"Invalid value for --max-requests: '{}'.", arg);
                return parsed;
            }

            parsed.maxRequestsBeforeExit = value;
            continue;
        }

        if (arg.rfind(L"--disconnect-after-batches=", 0) == 0)
        {
            uint32_t value = 0u;
            if (! TryParseUnsigned(arg.substr(std::wstring_view(L"--disconnect-after-batches=").size()), value))
            {
                parsed.errorMessage = std::format(L"Invalid value for --disconnect-after-batches: '{}'.", arg);
                return parsed;
            }

            parsed.disconnectAfterBatches = value;
            continue;
        }

        parsed.errorMessage = std::format(L"Unknown option '{}'.", arg);
        return parsed;
    }

    if (parsed.action == StartupAction::CompactStore && ! parsed.storeBackendExplicit && parsed.sqliteDatabasePath.empty())
    {
        parsed.persistentStoreKind = LocalSearchIndexCore::PersistentStoreKind::Sqlite;
    }

    const bool hasPipeOverrideOnly = ! parsed.pipeName.empty();
    const bool hasForegroundOnlyOptions =
        parsed.protocolVersion != SearchServiceBroker::kProtocolVersion || parsed.maxRequestsBeforeExit != 0u || parsed.disconnectAfterBatches != 0u;
    if (parsed.action != StartupAction::ShowHelp && hasForegroundOnlyOptions && parsed.action != StartupAction::RunForeground)
    {
        parsed.errorMessage = L"--protocol-version, --max-requests, and --disconnect-after-batches require --run-foreground.";
    }
    if (parsed.errorMessage.empty() && parsed.action != StartupAction::ShowHelp && hasPipeOverrideOnly && parsed.action != StartupAction::RunForeground &&
        parsed.action != StartupAction::RequestCompactService)
    {
        parsed.errorMessage = L"--pipe-name requires --run-foreground or --request-compact.";
    }

    const bool hasStoreOverrideOptions = ! parsed.storageRootDirectory.empty() || ! parsed.sqliteDatabasePath.empty() || parsed.storeBackendExplicit;
    if (parsed.errorMessage.empty() && parsed.action != StartupAction::ShowHelp && hasStoreOverrideOptions && parsed.action != StartupAction::RunForeground &&
        parsed.action != StartupAction::CompactStore)
    {
        parsed.errorMessage = L"--storage-root, --store-backend, and --sqlite-path require --run-foreground or --compact.";
    }

    return parsed;
}

DWORD WINAPI ServiceControlHandlerEx(DWORD control, DWORD eventType, void* eventData, void* context) noexcept
{
    UNREFERENCED_PARAMETER(eventType);
    UNREFERENCED_PARAMETER(eventData);
    UNREFERENCED_PARAMETER(context);

    if ((control == SERVICE_CONTROL_STOP || control == SERVICE_CONTROL_SHUTDOWN) && g_serviceStopEvent)
    {
        UpdateServiceStatus(SERVICE_STOP_PENDING, NO_ERROR, 5000u);
        g_serviceStopEvent.SetEvent();
        return NO_ERROR;
    }

    return ERROR_CALL_NOT_IMPLEMENTED;
}

void WINAPI ServiceMain(DWORD argc, wchar_t** argv) noexcept
{
    UNREFERENCED_PARAMETER(argc);
    UNREFERENCED_PARAMETER(argv);

    g_serviceStatusHandle = ::RegisterServiceCtrlHandlerExW(SearchServiceBroker::kServiceName, &ServiceControlHandlerEx, nullptr);
    if (! g_serviceStatusHandle)
    {
        return;
    }

    UpdateServiceStatus(SERVICE_START_PENDING, NO_ERROR, 5000u);

    wil::unique_handle instanceGuard;
    bool alreadyRunning      = false;
    const HRESULT instanceHr = AcquireSingleInstanceGuard(instanceGuard, alreadyRunning);
    if (FAILED(instanceHr))
    {
        UpdateServiceStatus(SERVICE_STOPPED, static_cast<DWORD>(HRESULT_CODE(instanceHr)), 0u);
        return;
    }
    if (alreadyRunning)
    {
        UpdateServiceStatus(SERVICE_STOPPED, ERROR_SERVICE_ALREADY_RUNNING, 0u);
        return;
    }

    g_serviceStopEvent.create();

    SearchServiceBroker::ServerOptions options{};
    const HRESULT hr = SearchServiceBroker::RunServer(options, g_serviceStopEvent.get(), nullptr);
    UpdateServiceStatus(SERVICE_STOPPED, FAILED(hr) ? static_cast<DWORD>(HRESULT_CODE(hr)) : NO_ERROR, 0u);
}
} // namespace

int wmain()
{
    const ParsedArguments parsed = ParseArguments();
    if (! parsed.errorMessage.empty())
    {
        PrintUsageError(parsed.errorMessage);
        return 1;
    }

    switch (parsed.action)
    {
        case StartupAction::ShowHelp: PrintHelpText(); return 0;

        case StartupAction::CompactStore:
        {
            const CommandResult result = RunCompactStore(parsed);
            WriteConsoleText(std::format(L"{}\r\n", result.message), FAILED(result.hr));
            return FAILED(result.hr) ? 1 : 0;
        }

        case StartupAction::RequestCompactService:
        {
            const CommandResult result = RequestServiceCompaction(parsed);
            WriteConsoleText(std::format(L"{}\r\n", result.message), FAILED(result.hr));
            return FAILED(result.hr) ? 1 : 0;
        }

        case StartupAction::RegisterService:
        {
            const CommandResult result = RegisterCurrentService();
            WriteConsoleText(std::format(L"{}\r\n", result.message), FAILED(result.hr));
            return FAILED(result.hr) ? 1 : 0;
        }

        case StartupAction::UnregisterService:
        {
            const CommandResult result = UnregisterCurrentService();
            WriteConsoleText(std::format(L"{}\r\n", result.message), FAILED(result.hr));
            return FAILED(result.hr) ? 1 : 0;
        }

        case StartupAction::RunForeground:
        {
            const CommandResult result = RunForegroundService(parsed);
            WriteConsoleText(std::format(L"{}\r\n", result.message), FAILED(result.hr));
            return FAILED(result.hr) ? 1 : 0;
        }

        case StartupAction::ServiceDispatcher:
        default: break;
    }

    SERVICE_TABLE_ENTRYW table[] = {
        {const_cast<LPWSTR>(SearchServiceBroker::kServiceName), &ServiceMain},
        {nullptr, nullptr},
    };

    if (::StartServiceCtrlDispatcherW(table) == 0)
    {
        const DWORD error = ::GetLastError();
        if (error == ERROR_FAILED_SERVICE_CONTROLLER_CONNECT)
        {
            WriteConsoleText(L"This executable is a Windows service. Use --run-foreground to run it in the current terminal or --help for usage.\r\n", true);
        }
        return 1;
    }

    return 0;
}
