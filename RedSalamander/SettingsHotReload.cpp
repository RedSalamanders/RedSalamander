#include "SettingsHotReload.h"

#include <algorithm>
#include <atomic>
#include <chrono>
#include <condition_variable>
#include <deque>
#include <exception>
#include <memory>
#include <mutex>
#include <string>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"
#include "HostServices.h"
#include "SettingsSave.h"
#include "SettingsSchemaExport.h"
#include "WindowBackdropPolicy.h"
#include "WindowMessages.h"
#include "resource.h"

namespace
{
using unique_change_notification = wil::unique_any<HANDLE, decltype(&::FindCloseChangeNotification), ::FindCloseChangeNotification>;

struct HotReloadState
{
    std::mutex mutex;
    HWND targetWindow = nullptr;
    std::wstring appId;
    std::filesystem::path settingsPath;
    std::filesystem::path settingsDirectory;
    wil::unique_event_nothrow stopEvent;
    wil::unique_event_nothrow readyEvent;
    std::jthread watchThread;
    std::unordered_set<HWND> participants;
    std::optional<Common::Settings::SettingsFileStamp> lastAppliedStamp;
    std::optional<Common::Settings::SettingsFileStamp> lastRejectedStamp;
    uint64_t sessionGeneration = 0;
    uint64_t internalSaveEpoch = 0;
    uint32_t internalSaveDepth = 0;
    bool internalSaveNotificationDeferred = false;
    bool invalidAlertVisible = false;

    HotReloadState()                                 = default;
    HotReloadState(const HotReloadState&)            = delete;
    HotReloadState& operator=(const HotReloadState&) = delete;
    HotReloadState(HotReloadState&&)                 = delete;
    HotReloadState& operator=(HotReloadState&&)      = delete;
};

// A detached persistence worker can outlive the bounded UI shutdown wait. Keep the watcher/session
// state alive until ExitProcess so a late worker cannot race a C++ static destructor. The allocation
// is intentionally process-lifetime and is reclaimed by the operating system.
HotReloadState& g_state = []() -> HotReloadState&
{
    auto state = std::make_unique<HotReloadState>();
    return *state.release();
}();
int g_invalidAlertCookieStorage = 0;

#ifdef ENABLE_TESTS
std::atomic<uint32_t> g_debugChangeNotificationOpenFailureCount{0u};
std::atomic<DWORD> g_debugChangeNotificationOpenFailureLastError{ERROR_PATH_NOT_FOUND};
std::atomic<DWORD> g_debugSettingsSavePostWriteDelayMs{0u};
std::atomic<DWORD> g_debugSettingsReloadPostStampDelayMs{0u};
std::atomic_bool g_debugSettingsReloadPostStampDelayActive{false};

[[nodiscard]] bool ConsumeDebugChangeNotificationOpenFailure(DWORD& lastError) noexcept
{
    uint32_t remaining = g_debugChangeNotificationOpenFailureCount.load(std::memory_order_acquire);
    while (remaining != 0u)
    {
        if (g_debugChangeNotificationOpenFailureCount.compare_exchange_weak(
                remaining, remaining - 1u, std::memory_order_acq_rel, std::memory_order_acquire))
        {
            lastError = g_debugChangeNotificationOpenFailureLastError.load(std::memory_order_acquire);
            if (lastError == ERROR_SUCCESS)
            {
                lastError = ERROR_GEN_FAILURE;
            }
            return true;
        }
    }

    return false;
}
#endif

[[nodiscard]] void* InvalidAlertCookie() noexcept
{
    return &g_invalidAlertCookieStorage;
}

[[nodiscard]] bool ShouldIgnoreStampLocked(const Common::Settings::SettingsFileStamp& stamp) noexcept
{
    return (g_state.lastAppliedStamp.has_value() && g_state.lastAppliedStamp.value() == stamp) ||
           (g_state.lastRejectedStamp.has_value() && g_state.lastRejectedStamp.value() == stamp);
}

void PostSettingsFileChanged(HWND targetWindow) noexcept
{
    if (! targetWindow)
    {
        return;
    }

    auto payload       = std::make_unique<SettingsHotReload::SettingsFileChangedPayload>();
    payload->tickCount = GetTickCount64();
    if (! PostMessagePayload(targetWindow, WndMsg::kSettingsFileChanged, 0, std::move(payload)))
    {
        const DWORD lastError = GetLastError();
        if (lastError != ERROR_INVALID_WINDOW_HANDLE && lastError != ERROR_SUCCESS)
        {
            Debug::Warning(L"SettingsHotReload: failed to post settings change notification (gle=0x{:08X})", lastError);
        }
    }
}

struct InternalSaveToken final
{
    uint64_t sessionGeneration = 0;
    bool active                = false;
};

[[nodiscard]] InternalSaveToken BeginInternalSave(std::wstring_view appId) noexcept
{
    std::scoped_lock lock(g_state.mutex);
    if (g_state.appId.empty() || ! OrdinalString::EqualsNoCase(g_state.appId, appId))
    {
        return {};
    }

    ++g_state.internalSaveDepth;
    ++g_state.internalSaveEpoch;
    return InternalSaveToken{.sessionGeneration = g_state.sessionGeneration, .active = true};
}

void CompleteInternalSave(std::wstring_view appId,
                          InternalSaveToken token,
                          const Common::Settings::SettingsFileStamp* appliedStamp) noexcept
{
    if (! token.active)
    {
        return;
    }

    HWND deferredTarget = nullptr;
    {
        std::scoped_lock lock(g_state.mutex);
        if (token.sessionGeneration != g_state.sessionGeneration || g_state.internalSaveDepth == 0u ||
            ! OrdinalString::EqualsNoCase(g_state.appId, appId))
        {
            return;
        }

        if (appliedStamp)
        {
            g_state.lastAppliedStamp = *appliedStamp;
            g_state.lastRejectedStamp.reset();
        }
        --g_state.internalSaveDepth;
        ++g_state.internalSaveEpoch;
        if (g_state.internalSaveDepth == 0u && g_state.internalSaveNotificationDeferred)
        {
            g_state.internalSaveNotificationDeferred = false;
            deferredTarget                           = g_state.targetWindow;
        }
    }

    PostSettingsFileChanged(deferredTarget);
}

[[nodiscard]] bool DeferNotificationForInternalSaveLocked() noexcept
{
    if (g_state.internalSaveDepth == 0u)
    {
        return false;
    }

    g_state.internalSaveNotificationDeferred = true;
    return true;
}

void WatchSettingsDirectoryThread(HWND targetWindow,
                                  HANDLE stopEventHandle,
                                  HANDLE readyEventHandle,
                                  std::filesystem::path directoryPath,
                                  std::wstring appId,
                                  std::optional<Common::Settings::SettingsFileStamp> initialStamp) noexcept
{
    if (! targetWindow || directoryPath.empty())
    {
        return;
    }

    // Readiness means the watcher is alive and will keep retrying; it does not require the
    // directory/change-notification handle to exist yet.
    if (readyEventHandle)
    {
        static_cast<void>(SetEvent(readyEventHandle));
    }

    const std::wstring directoryText = directoryPath.wstring();
    constexpr DWORD kNotifyFilter    = FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_LAST_WRITE | FILE_NOTIFY_CHANGE_SIZE;

    unique_change_notification changeNotification;
    bool catchUpPending = true;

    for (;;)
    {
        if (! changeNotification)
        {
            std::error_code directoryError;
            static_cast<void>(std::filesystem::create_directories(directoryPath, directoryError));
            if (directoryError)
            {
                Debug::Warning(L"SettingsHotReload: settings directory is unavailable; watcher will retry '{}' (error={})",
                               directoryText,
                               directoryError.value());
                if (WaitForSingleObject(stopEventHandle, 1000) == WAIT_OBJECT_0)
                {
                    return;
                }
                continue;
            }

            HANDLE raw = nullptr;
#ifdef ENABLE_TESTS
            DWORD forcedLastError = ERROR_SUCCESS;
            if (ConsumeDebugChangeNotificationOpenFailure(forcedLastError))
            {
                SetLastError(forcedLastError);
                raw = INVALID_HANDLE_VALUE;
            }
            else
#endif
            {
                raw = FindFirstChangeNotificationW(directoryText.c_str(), FALSE, kNotifyFilter);
            }
            if (raw == nullptr || raw == INVALID_HANDLE_VALUE)
            {
                static_cast<void>(Debug::ErrorWithLastError(L"SettingsHotReload: FindFirstChangeNotificationW failed for '{}'", directoryText));
                if (WaitForSingleObject(stopEventHandle, 1000) == WAIT_OBJECT_0)
                {
                    return;
                }
                continue;
            }

            changeNotification.reset(raw);
            if (catchUpPending)
            {
                Common::Settings::SettingsFileStamp armedStamp{};
                const HRESULT stampHr = Common::Settings::TryGetSettingsFileStamp(appId, armedStamp);
                if (stampHr == S_OK)
                {
                    if (! initialStamp.has_value() || initialStamp.value() != armedStamp)
                    {
                        PostSettingsFileChanged(targetWindow);
                    }
                }
                else if (stampHr == S_FALSE)
                {
                    if (initialStamp.has_value())
                    {
                        PostSettingsFileChanged(targetWindow);
                    }
                }
                else
                {
                    Debug::Warning(L"SettingsHotReload: catch-up stamp query failed (hr=0x{:08X})", static_cast<unsigned long>(stampHr));
                    PostSettingsFileChanged(targetWindow);
                }
                catchUpPending = false;
            }
        }

        HANDLE waitHandles[2]  = {stopEventHandle, changeNotification.get()};
        const DWORD waitResult = WaitForMultipleObjects(2u, waitHandles, FALSE, INFINITE);
        if (waitResult == WAIT_OBJECT_0)
        {
            return;
        }

        if (waitResult == (WAIT_OBJECT_0 + 1))
        {
            PostSettingsFileChanged(targetWindow);

            if (! FindNextChangeNotification(changeNotification.get()))
            {
                static_cast<void>(Debug::ErrorWithLastError(L"SettingsHotReload: FindNextChangeNotification failed for '{}'", directoryText));
                changeNotification.reset();
            }

            continue;
        }

        if (waitResult == WAIT_FAILED)
        {
            const DWORD lastError = GetLastError();
            Debug::Error(L"SettingsHotReload: WaitForMultipleObjects failed (gle=0x{:08X})", lastError);
        }
        return;
    }
}

enum class SettingsSaveShutdownState : uint8_t
{
    Running,
    FinalSavePending,
    FinalSaveQueued,
    ShuttingDown,
};

[[nodiscard]] bool IsProcessShutdownStarted(const std::atomic<SettingsSaveShutdownState>& state) noexcept
{
    return state.load(std::memory_order_acquire) != SettingsSaveShutdownState::Running;
}

[[nodiscard]] HRESULT SavePreparedSettingsAndSchema(std::wstring_view appId,
                                                    Common::Settings::Settings& settings,
                                                    std::span<const PluginConfigurationSchemaSource> pluginSchemas,
                                                    bool writeSchema,
                                                    const std::atomic<SettingsSaveShutdownState>& shutdownState,
                                                    bool allowDuringProcessShutdown) noexcept
{
    if (appId.empty())
    {
        return E_INVALIDARG;
    }
    if (settings.persistence.savePermission != Common::Settings::SettingsSavePermission::Automatic)
    {
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }

    const bool shutdownAtStart = IsProcessShutdownStarted(shutdownState);
    if (shutdownAtStart && ! allowDuringProcessShutdown)
    {
        return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
    }

    const InternalSaveToken internalSave = shutdownAtStart ? InternalSaveToken{} : BeginInternalSave(appId);
    bool internalSaveCompleted           = false;
    const auto endInternalSave = wil::scope_exit([&]() noexcept
    {
        if (! internalSaveCompleted && ! IsProcessShutdownStarted(shutdownState))
        {
            CompleteInternalSave(appId, internalSave, nullptr);
        }
    });

    Common::Settings::Settings settingsToSave = SettingsSave::PrepareForSave(settings);
    Common::Settings::SettingsFileStamp writtenStamp{};
    const HRESULT saveHr = Common::Settings::SaveSettingsValuesOnlyWithStamp(appId, settingsToSave, writtenStamp);
    if (FAILED(saveHr))
    {
        return saveHr;
    }
    settings.persistence.expectedFileStamp = writtenStamp;

    // Publish the identity of the exact atomic replacement and end suppression before schema I/O.
    // A later external replacement therefore has a different stamp and cannot be mislabeled as ours.
    if (internalSave.active && ! IsProcessShutdownStarted(shutdownState))
    {
        CompleteInternalSave(appId, internalSave, &writtenStamp);
        internalSaveCompleted = true;
    }

#ifdef ENABLE_TESTS
    const DWORD postWriteDelayMs = IsProcessShutdownStarted(shutdownState)
                                       ? 0u
                                       : g_debugSettingsSavePostWriteDelayMs.load(std::memory_order_acquire);
    if (postWriteDelayMs != 0u)
    {
        Sleep(postWriteDelayMs);
    }
#endif

    if (writeSchema && (! IsProcessShutdownStarted(shutdownState) || allowDuringProcessShutdown))
    {
        const HRESULT schemaHr = SaveAggregatedSettingsSchema(appId, pluginSchemas);
        if (FAILED(schemaHr) && ! IsProcessShutdownStarted(shutdownState))
        {
            Debug::Error(L"SettingsHotReload: SaveAggregatedSettingsSchema failed (hr=0x{:08X})", static_cast<unsigned long>(schemaHr));
        }
    }

    return saveHr;
}

constexpr DWORD kSettingsSaveShutdownFlushTimeoutMs = 5000u;

class SerializedSettingsSaveCoordinator final
{
public:
    struct SynchronousSaveOptions final
    {
        bool writeSchema          = true;
        bool beginProcessShutdown = false;
    };

    struct Completion final
    {
        Completion()                             = default;
        Completion(const Completion&)            = delete;
        Completion(Completion&&)                 = delete;
        Completion& operator=(const Completion&) = delete;
        Completion& operator=(Completion&&)      = delete;

        std::mutex mutex;
        std::condition_variable cv;
        bool done  = false;
        HRESULT hr = E_PENDING;
        std::optional<Common::Settings::SettingsFileStamp> committedStamp;
    };

    struct Request final
    {
        std::wstring appId;
        Common::Settings::Settings settings;
        std::vector<PluginConfigurationSchemaSource> pluginSchemas;
        std::wstring telemetryMetric;
        std::wstring telemetryContext;
        std::shared_ptr<Completion> completion;
        std::chrono::steady_clock::time_point readyAt{};
        uint64_t generation = 0;
        bool writeSchema    = true;
        bool asynchronous   = false;
        bool allowDuringProcessShutdown = false;
        bool processFinalSave            = false;
    };

    SerializedSettingsSaveCoordinator() = default;

    [[nodiscard]] static std::shared_ptr<SerializedSettingsSaveCoordinator> CreateProcessLifetime()
    {
        auto coordinator = std::make_shared<SerializedSettingsSaveCoordinator>();

        // This coordinator belongs to RedSalamander.exe, not a plugin DLL. The worker deliberately
        // owns the coordinator for the rest of the process so a storage call that outlives the
        // bounded UI shutdown flush cannot reappear as an unbounded static-destructor join.
        std::thread worker([lifetime = coordinator]() noexcept { lifetime->ThreadMain(); });
        worker.detach();
        return coordinator;
    }

    SerializedSettingsSaveCoordinator(const SerializedSettingsSaveCoordinator&)            = delete;
    SerializedSettingsSaveCoordinator(SerializedSettingsSaveCoordinator&&)                 = delete;
    SerializedSettingsSaveCoordinator& operator=(const SerializedSettingsSaveCoordinator&) = delete;
    SerializedSettingsSaveCoordinator& operator=(SerializedSettingsSaveCoordinator&&)      = delete;

    HRESULT EnqueueAsync(std::wstring_view appId,
                         const Common::Settings::Settings& settings,
                         std::wstring_view telemetryMetric,
                         std::wstring_view telemetryContext)
    {
        if (appId.empty())
        {
            return E_INVALIDARG;
        }
        if (settings.persistence.savePermission != Common::Settings::SettingsSavePermission::Automatic)
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }

        std::unique_lock submissionLock(_submissionMutex);
        if (_shutdownState.load(std::memory_order_acquire) != SettingsSaveShutdownState::Running)
        {
            return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
        }
        Request request{};
        request.appId            = appId;
        request.settings         = settings;
        request.telemetryMetric  = telemetryMetric;
        request.telemetryContext = telemetryContext;
        request.writeSchema      = false;
        request.asynchronous     = true;

        {
            std::scoped_lock lock(_mutex);
            request.generation = ++_lastQueuedGeneration;
            request.readyAt     = std::chrono::steady_clock::now() + kDebounceDelay;
            requestThreadId();
            if (! _requests.empty() && _requests.back().asynchronous && _requests.back().appId == request.appId)
            {
                _requests.back() = std::move(request);
                ++_totalCoalescedCount;
            }
            else
            {
                _requests.push_back(std::move(request));
            }
        }
        _cv.notify_all();
        return S_OK;
    }

    HRESULT SaveSynchronously(std::wstring_view appId,
                              Common::Settings::Settings& settings,
                              std::span<const PluginConfigurationSchemaSource> pluginSchemas,
                              DWORD timeoutMs,
                              const SynchronousSaveOptions options)
    {
        if (appId.empty())
        {
            return E_INVALIDARG;
        }
        if (settings.persistence.savePermission != Common::Settings::SettingsSavePermission::Automatic)
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }

        std::shared_ptr<Completion> completion;
        bool queued = false;
        {
            std::unique_lock submissionLock(_submissionMutex);
            const SettingsSaveShutdownState shutdownState = _shutdownState.load(std::memory_order_acquire);
            if (! options.beginProcessShutdown && shutdownState != SettingsSaveShutdownState::Running)
            {
                return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
            }

            if (options.beginProcessShutdown &&
                (shutdownState == SettingsSaveShutdownState::FinalSaveQueued || shutdownState == SettingsSaveShutdownState::ShuttingDown))
            {
                completion = _finalSaveCompletion;
            }
            else
            {
                completion = std::make_shared<Completion>();
                Request request{};
                request.appId    = appId;
                request.settings = settings;
                if (options.writeSchema)
                {
                    request.pluginSchemas = pluginSchemas.empty()
                                                ? CollectPluginConfigurationSchemas(request.settings)
                                                : std::vector<PluginConfigurationSchemaSource>(pluginSchemas.begin(), pluginSchemas.end());
                }
                request.completion                 = completion;
                request.readyAt                    = std::chrono::steady_clock::now();
                request.writeSchema                = options.writeSchema;
                request.asynchronous               = false;
                request.allowDuringProcessShutdown = options.beginProcessShutdown;
                request.processFinalSave            = options.beginProcessShutdown;

                {
                    std::scoped_lock lock(_mutex);
                    request.generation = ++_lastQueuedGeneration;
                    requestThreadId();
                    _requests.push_back(std::move(request));
                    if (options.beginProcessShutdown)
                    {
                        // The worker cannot dequeue an older request between publication of this
                        // final request and its shutdown fence. Older requests therefore fail
                        // closed unless one was already saving, in which case this final write
                        // remains ordered after it and wins.
                        _forceFlush = true;
                        _shutdownState.store(SettingsSaveShutdownState::FinalSaveQueued, std::memory_order_release);
                    }
                }
                if (options.beginProcessShutdown)
                {
                    _finalSaveCompletion = completion;
                }
                queued = true;
            }
        }

        if (! completion)
        {
            return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
        }
        if (queued)
        {
            _cv.notify_all();
        }

        std::unique_lock completionLock(completion->mutex);
        const auto completed = [&]() noexcept { return completion->done; };
        bool finished = false;
        if (timeoutMs == INFINITE)
        {
            completion->cv.wait(completionLock, completed);
            finished = true;
        }
        else
        {
            finished = completion->cv.wait_for(completionLock, std::chrono::milliseconds(timeoutMs), completed);
        }
        if (! finished)
        {
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }
        const HRESULT result = completion->hr;
        if (SUCCEEDED(result) && completion->committedStamp.has_value())
        {
            settings.persistence.expectedFileStamp = completion->committedStamp;
        }
        return result;
    }

    void BeginProcessShutdown() noexcept
    {
        std::scoped_lock submissionLock(_submissionMutex);
        if (_shutdownState.load(std::memory_order_acquire) == SettingsSaveShutdownState::Running)
        {
            _shutdownState.store(SettingsSaveShutdownState::FinalSavePending, std::memory_order_release);
        }
        _cv.notify_all();
    }

    bool FlushFor(DWORD timeoutMs) noexcept
    {
        std::unique_lock submissionLock(_submissionMutex);
        std::unique_lock lock(_mutex);
        if (_requests.empty() && ! _saveInProgress)
        {
            _forceFlush = false;
            return true;
        }

        _forceFlush = true;
        _cv.notify_all();
        submissionLock.unlock();
        const auto complete = [&]() noexcept { return _requests.empty() && ! _saveInProgress; };
        bool flushed = false;
        if (timeoutMs == INFINITE)
        {
            _cv.wait(lock, complete);
            flushed = true;
        }
        else
        {
            flushed = _cv.wait_for(lock, std::chrono::milliseconds(timeoutMs), complete);
        }
        if (flushed)
        {
            _forceFlush = false;
        }
        return flushed;
    }

#ifdef ENABLE_TESTS
    SettingsHotReload::SettingsSaveDebugSnapshot GetDebugSnapshot() noexcept
    {
        std::scoped_lock lock(_mutex);
        SettingsHotReload::SettingsSaveDebugSnapshot snapshot{};
        snapshot.queuedGeneration    = _lastQueuedGeneration;
        snapshot.completedGeneration = _lastCompletedGeneration;
        snapshot.coalescedCount      = _totalCoalescedCount;
        snapshot.lastQueueThreadId   = _lastQueueThreadId;
        snapshot.lastSaveThreadId    = _lastSaveThreadId;
        snapshot.pending             = ! _requests.empty();
        snapshot.saveInProgress      = _saveInProgress;
        return snapshot;
    }

    void SetDebugSaveDelay(DWORD delayMs) noexcept
    {
        std::scoped_lock lock(_mutex);
        _debugSaveDelayMs = delayMs;
    }
#endif

private:
    struct StampLineage final
    {
        std::optional<Common::Settings::SettingsFileStamp> source;
        std::optional<Common::Settings::SettingsFileStamp> committed;
    };

    void AdvanceExpectedStampLocked(Request& request) noexcept
    {
        const auto lineage = _stampLineage.find(request.appId);
        if (lineage == _stampLineage.end())
        {
            return;
        }
        const auto& expected = request.settings.persistence.expectedFileStamp;
        if (expected == lineage->second.source || expected == lineage->second.committed)
        {
            request.settings.persistence.expectedFileStamp = lineage->second.committed;
        }
    }

    void requestThreadId() noexcept
    {
        _lastQueueThreadId = GetCurrentThreadId();
    }

    void ThreadMain() noexcept
    {
        for (;;)
        {
            std::optional<Request> request;
            uint64_t coalescedCount = 0;
#ifdef ENABLE_TESTS
            DWORD debugSaveDelayMs = 0;
#endif
            {
                std::unique_lock lock(_mutex);
                _cv.wait(lock, [&]() noexcept { return ! _requests.empty(); });

                if (! _forceFlush && ! _requests.empty() && _requests.front().asynchronous)
                {
                    const uint64_t generation = _requests.front().generation;
                    const auto readyAt        = _requests.front().readyAt;
                    const bool interrupted = _cv.wait_until(lock, readyAt, [&]() noexcept
                    { return _forceFlush || _requests.empty() || _requests.front().generation != generation; });
                    if (interrupted && ! _forceFlush)
                    {
                        continue;
                    }
                }

                if (_requests.empty())
                {
                    continue;
                }

                AdvanceExpectedStampLocked(_requests.front());
                request = std::move(_requests.front());
                _requests.pop_front();
                _saveInProgress   = true;
                _lastSaveThreadId = GetCurrentThreadId();
                coalescedCount    = _totalCoalescedCount;
#ifdef ENABLE_TESTS
                debugSaveDelayMs = _debugSaveDelayMs;
#endif
            }

            Request& current = request.value();
            const std::optional<Common::Settings::SettingsFileStamp> saveExpectedStamp = current.settings.persistence.expectedFileStamp;
#ifdef ENABLE_TESTS
            if (debugSaveDelayMs != 0)
            {
                Sleep(debugSaveDelayMs);
            }
#endif
            const uint64_t saveStart = static_cast<uint64_t>(
                std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now().time_since_epoch()).count());
            const HRESULT saveHr = SavePreparedSettingsAndSchema(current.appId,
                                                                 current.settings,
                                                                 current.pluginSchemas,
                                                                 current.writeSchema,
                                                                 _shutdownState,
                                                                 current.allowDuringProcessShutdown);
            const uint64_t saveEnd = static_cast<uint64_t>(
                std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now().time_since_epoch()).count());
            const bool processShuttingDown = IsProcessShutdownStarted(_shutdownState);
            if (! processShuttingDown && current.asynchronous && FAILED(saveHr))
            {
                Debug::Error(L"SettingsHotReload: asynchronous settings save failed for '{}' ({}) hr=0x{:08X}",
                             current.appId,
                             current.telemetryContext,
                             static_cast<unsigned long>(saveHr));
            }
            if (! processShuttingDown && ! current.telemetryMetric.empty() && Debug::Perf::IsCaptureEnabled())
            {
                Debug::Perf::Emit(current.telemetryMetric,
                                  current.telemetryContext,
                                  saveEnd >= saveStart ? saveEnd - saveStart : 0u,
                                  current.generation,
                                  coalescedCount,
                                  saveHr);
            }

            if (current.completion)
            {
                {
                    std::scoped_lock completionLock(current.completion->mutex);
                    current.completion->hr = saveHr;
                    if (SUCCEEDED(saveHr))
                    {
                        current.completion->committedStamp = current.settings.persistence.expectedFileStamp;
                    }
                    current.completion->done = true;
                }
                current.completion->cv.notify_all();
            }

            if (current.processFinalSave)
            {
                std::scoped_lock submissionLock(_submissionMutex);
                _shutdownState.store(SettingsSaveShutdownState::ShuttingDown, std::memory_order_release);
            }

            {
                std::scoped_lock lock(_mutex);
                if (SUCCEEDED(saveHr))
                {
                    auto lineage = _stampLineage.find(current.appId);
                    if (lineage == _stampLineage.end())
                    {
                        _stampLineage.emplace(current.appId,
                                              StampLineage{.source = saveExpectedStamp,
                                                           .committed = current.settings.persistence.expectedFileStamp});
                    }
                    else
                    {
                        if (saveExpectedStamp != lineage->second.source && saveExpectedStamp != lineage->second.committed)
                        {
                            lineage->second.source = saveExpectedStamp;
                        }
                        lineage->second.committed = current.settings.persistence.expectedFileStamp;
                    }
                }
                _lastCompletedGeneration = current.generation;
                _saveInProgress          = false;
            }
            _cv.notify_all();
        }
    }

    static constexpr std::chrono::milliseconds kDebounceDelay{200};

    std::mutex _mutex;
    std::mutex _submissionMutex;
    std::condition_variable _cv;
    std::deque<Request> _requests;
    std::unordered_map<std::wstring, StampLineage> _stampLineage;
    uint64_t _lastQueuedGeneration    = 0;
    uint64_t _lastCompletedGeneration = 0;
    uint64_t _totalCoalescedCount      = 0;
    DWORD _lastQueueThreadId           = 0;
    DWORD _lastSaveThreadId            = 0;
    bool _saveInProgress               = false;
    bool _forceFlush                   = false;
    std::atomic<SettingsSaveShutdownState> _shutdownState{SettingsSaveShutdownState::Running};
    std::shared_ptr<Completion> _finalSaveCompletion;
#ifdef ENABLE_TESTS
    DWORD _debugSaveDelayMs = 0;
#endif
};

SerializedSettingsSaveCoordinator& GetSettingsSaveCoordinator()
{
    static const std::shared_ptr<SerializedSettingsSaveCoordinator> coordinator =
        SerializedSettingsSaveCoordinator::CreateProcessLifetime();
    return *coordinator;
}

[[nodiscard]] std::wstring BuildExternalReloadConflictMessage(std::wstring_view editorName) noexcept
{
    return FormatStringResource(nullptr, IDS_FMT_SETTINGS_RELOAD_CONFLICT_KEEP_EDITING, editorName);
}

[[nodiscard]] std::wstring BuildStaleSaveConflictMessage(std::wstring_view editorName) noexcept
{
    return FormatStringResource(nullptr, IDS_FMT_SETTINGS_RELOAD_CONFLICT_STALE_SAVE, editorName);
}

[[nodiscard]] Common::Settings::UiSettings GetUiSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    if (settings.ui.has_value())
    {
        return settings.ui.value();
    }
    return {};
}

[[nodiscard]] AppBackdropType ToAppBackdropType(Common::WindowBackdrop::Kind kind) noexcept
{
    switch (kind)
    {
        case Common::WindowBackdrop::Kind::Mica: return AppBackdropType::Mica;
        case Common::WindowBackdrop::Kind::Acrylic: return AppBackdropType::Acrylic;
        case Common::WindowBackdrop::Kind::MicaAlt: return AppBackdropType::MicaAlt;
        case Common::WindowBackdrop::Kind::None:
        default: return AppBackdropType::None;
    }
}

void ApplyResolvedWindowBackdrop(Common::Settings::WindowBackdropMode mode, AppTheme& theme) noexcept
{
    theme.primaryWindowBackdrop = ToAppBackdropType(Common::WindowBackdrop::Resolve(mode, Common::WindowBackdrop::Target::Primary, false));
    theme.toolWindowBackdrop    = ToAppBackdropType(Common::WindowBackdrop::Resolve(mode, Common::WindowBackdrop::Target::Tool, false));

    if (! theme.highContrast && (theme.primaryWindowBackdrop != AppBackdropType::None || theme.toolWindowBackdrop != AppBackdropType::None))
    {
        theme.titleBar.captionColor.reset();
        theme.titleBar.borderColor.reset();
        theme.titleBar.textColor.reset();
    }
}

[[nodiscard]] bool IsRetryableSettingsLoadFailure(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) || hr == HRESULT_FROM_WIN32(ERROR_LOCK_VIOLATION) || hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
}

[[nodiscard]] bool IsInvalidSettingsLoadFailure(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}
} // namespace

namespace SettingsHotReload
{
void ApplyUiPreferencesToTheme(const Common::Settings::Settings& settings, AppTheme& theme) noexcept
{
    const Common::Settings::UiSettings ui = GetUiSettingsOrDefault(settings);
    theme.compactMode                     = ui.compactMode;

    switch (ui.reducedMotion)
    {
        case Common::Settings::ReducedMotionMode::On: theme.reducedMotionOverride = true; break;
        case Common::Settings::ReducedMotionMode::Off: theme.reducedMotionOverride = false; break;
        case Common::Settings::ReducedMotionMode::System: theme.reducedMotionOverride.reset(); break;
    }

    ApplyResolvedWindowBackdrop(ui.windowBackdrop, theme);
}

#ifdef ENABLE_TESTS
void DebugSetChangeNotificationOpenFailuresForSelfTest(uint32_t failureCount, DWORD lastError) noexcept
{
    g_debugChangeNotificationOpenFailureLastError.store(lastError, std::memory_order_release);
    g_debugChangeNotificationOpenFailureCount.store(failureCount, std::memory_order_release);
}

void DebugSetSettingsSaveDelayForSelfTest(DWORD delayMs) noexcept
{
    try
    {
        GetSettingsSaveCoordinator().SetDebugSaveDelay(delayMs);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // noexcept test boundary: keep the harness alive and report the unavailable injection hook once.
        Debug::Error(L"SettingsHotReload: failed to configure the settings-save self-test delay.");
    }
}

void DebugSetSettingsSavePostWriteDelayForSelfTest(DWORD delayMs) noexcept
{
    g_debugSettingsSavePostWriteDelayMs.store(delayMs, std::memory_order_release);
}

void DebugSetSettingsReloadPostStampDelayForSelfTest(DWORD delayMs) noexcept
{
    g_debugSettingsReloadPostStampDelayActive.store(false, std::memory_order_release);
    g_debugSettingsReloadPostStampDelayMs.store(delayMs, std::memory_order_release);
}

bool DebugIsSettingsReloadPostStampDelayActiveForSelfTest() noexcept
{
    return g_debugSettingsReloadPostStampDelayActive.load(std::memory_order_acquire);
}

SettingsSaveDebugSnapshot DebugGetSettingsSaveSnapshotForSelfTest() noexcept
{
    try
    {
        return GetSettingsSaveCoordinator().GetDebugSnapshot();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // noexcept test boundary: return an empty snapshot after logging the unavailable diagnostics.
        Debug::Error(L"SettingsHotReload: failed to capture settings-save debug state.");
        return {};
    }
}
#endif

HRESULT Start(HWND targetWindow, std::wstring_view appId) noexcept
{
    Stop();

    if (! targetWindow || ! IsWindow(targetWindow) || appId.empty())
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(appId);
    if (settingsPath.empty())
    {
        return E_FAIL;
    }

    const std::filesystem::path settingsDirectory = settingsPath.parent_path();
    if (settingsDirectory.empty())
    {
        return E_FAIL;
    }

    wil::unique_event_nothrow stopEvent;
    stopEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! stopEvent)
    {
        const DWORD lastError = GetLastError();
        return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    wil::unique_event_nothrow readyEvent;
    readyEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! readyEvent)
    {
        const DWORD lastError = GetLastError();
        return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    std::optional<Common::Settings::SettingsFileStamp> initialStamp;
    Common::Settings::SettingsFileStamp stamp{};
    const HRESULT stampHr = Common::Settings::TryGetSettingsFileStamp(appId, stamp);
    if (stampHr == S_OK)
    {
        initialStamp = stamp;
    }
    else if (FAILED(stampHr))
    {
        Debug::Warning(L"SettingsHotReload: initial file stamp query failed (hr=0x{:08X})", static_cast<unsigned long>(stampHr));
    }

    {
        std::scoped_lock lock(g_state.mutex);
        g_state.targetWindow      = targetWindow;
        g_state.appId             = std::wstring(appId);
        g_state.settingsPath      = settingsPath;
        g_state.settingsDirectory = settingsDirectory;
        g_state.stopEvent         = std::move(stopEvent);
        g_state.readyEvent        = std::move(readyEvent);
        g_state.lastAppliedStamp  = initialStamp;
        g_state.lastRejectedStamp.reset();
        ++g_state.sessionGeneration;
        ++g_state.internalSaveEpoch;
        g_state.internalSaveDepth                = 0u;
        g_state.internalSaveNotificationDeferred = false;
        g_state.invalidAlertVisible = false;
    }

    HANDLE stopHandle = nullptr;
    HANDLE readyHandle = nullptr;
    {
        std::scoped_lock lock(g_state.mutex);
        stopHandle = g_state.stopEvent.get();
        readyHandle = g_state.readyEvent.get();
    }

    g_state.watchThread = std::jthread(
        [targetWindow, stopHandle, readyHandle, settingsDirectory, watchedAppId = std::wstring(appId), initialStamp](std::stop_token) noexcept
        { WatchSettingsDirectoryThread(targetWindow, stopHandle, readyHandle, settingsDirectory, watchedAppId, initialStamp); });

    return S_OK;
}

namespace
{
void StopWatcherCore(bool flushQueuedSaves) noexcept
{
    if (flushQueuedSaves && ! FlushQueuedSettingsSaves(kSettingsSaveShutdownFlushTimeoutMs))
    {
        Debug::Warning(L"SettingsHotReload: settings-save shutdown flush exceeded {} ms; the worker retains copied snapshots and will continue safely.",
                       kSettingsSaveShutdownFlushTimeoutMs);
    }

    ClearInvalidReloadAlert();

    {
        std::scoped_lock lock(g_state.mutex);
        if (g_state.stopEvent)
        {
            static_cast<void>(SetEvent(g_state.stopEvent.get()));
        }
    }

    if (g_state.watchThread.joinable())
    {
        g_state.watchThread.join();
    }

    std::scoped_lock lock(g_state.mutex);
    g_state.stopEvent.reset();
    g_state.readyEvent.reset();
    g_state.targetWindow = nullptr;
    g_state.appId.clear();
    g_state.settingsPath.clear();
    g_state.settingsDirectory.clear();
    g_state.participants.clear();
    g_state.lastAppliedStamp.reset();
    g_state.lastRejectedStamp.reset();
    ++g_state.sessionGeneration;
    ++g_state.internalSaveEpoch;
    g_state.internalSaveDepth                = 0u;
    g_state.internalSaveNotificationDeferred = false;
}
} // namespace

void Stop() noexcept
{
    StopWatcherCore(true);
}

void StopWatchingForProcessShutdown() noexcept
{
    StopWatcherCore(false);
}

void BeginProcessShutdown() noexcept
{
    try
    {
        GetSettingsSaveCoordinator().BeginProcessShutdown();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // If coordinator construction itself failed, there is no detached persistence worker to fence.
        Debug::Error(L"SettingsHotReload: failed to enter process-shutdown mode.");
    }
}

HRESULT SaveSettingsAndSchema(std::wstring_view appId, Common::Settings::Settings& settings) noexcept
{
    try
    {
        return GetSettingsSaveCoordinator().SaveSynchronously(
            appId, settings, {}, INFINITE, {.writeSchema = true, .beginProcessShutdown = false});
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // noexcept persistence boundary: report the failed serialized submission and return an HRESULT.
        Debug::Error(L"SettingsHotReload: failed to queue serialized settings save for '{}'", appId);
        return E_FAIL;
    }
}

HRESULT SaveSettingsAndSchema(std::wstring_view appId,
                              Common::Settings::Settings& settings,
                              std::span<const PluginConfigurationSchemaSource> pluginSchemas) noexcept
{
    try
    {
        return GetSettingsSaveCoordinator().SaveSynchronously(
            appId, settings, pluginSchemas, INFINITE, {.writeSchema = true, .beginProcessShutdown = false});
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // noexcept persistence boundary: report the failed serialized submission and return an HRESULT.
        Debug::Error(L"SettingsHotReload: failed to queue serialized settings/schema save for '{}'", appId);
        return E_FAIL;
    }
}

HRESULT SaveSettingsForSessionEnd(std::wstring_view appId,
                                  const Common::Settings::Settings& settings,
                                  const DWORD timeoutMs) noexcept
{
    try
    {
        // The regular synchronous API returns the committed file stamp through its mutable
        // snapshot. Session end has already captured an immutable snapshot, so keep that
        // caller-owned state unchanged while the coordinator owns this final copy.
        Common::Settings::Settings sessionEndSettings = settings;
        return GetSettingsSaveCoordinator().SaveSynchronously(
            appId, sessionEndSettings, {}, timeoutMs, {.writeSchema = false, .beginProcessShutdown = true});
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // noexcept session-end boundary: the caller must remain bounded even if submission setup fails.
        Debug::Error(L"SettingsHotReload: failed to queue the final session-end settings save for '{}'", appId);
        return E_FAIL;
    }
}

HRESULT SaveSettingsAndSchemaForProcessShutdown(std::wstring_view appId,
                                                Common::Settings::Settings& settings,
                                                DWORD timeoutMs) noexcept
{
    return SaveSettingsAndSchemaForProcessShutdown(appId, settings, {}, timeoutMs);
}

HRESULT SaveSettingsAndSchemaForProcessShutdown(std::wstring_view appId,
                                                Common::Settings::Settings& settings,
                                                std::span<const PluginConfigurationSchemaSource> pluginSchemas,
                                                DWORD timeoutMs) noexcept
{
    try
    {
        return GetSettingsSaveCoordinator().SaveSynchronously(
            appId, settings, pluginSchemas, timeoutMs, {.writeSchema = true, .beginProcessShutdown = true});
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // noexcept process-shutdown boundary: copied snapshots remain coordinator-owned on timeout/failure.
        Debug::Error(L"SettingsHotReload: failed to queue the final bounded settings/schema save for '{}'", appId);
        return E_FAIL;
    }
}

HRESULT ReplaceBlockedSettingsAndSchema(std::wstring_view appId,
                                        Common::Settings::Settings& settings,
                                        std::filesystem::path& backupPath) noexcept
{
    backupPath.clear();
    if (settings.persistence.savePermission != Common::Settings::SettingsSavePermission::ExplicitReplacementRequired)
    {
        return E_INVALIDARG;
    }

    const HRESULT backupHr = Common::Settings::BackupSettingsForExplicitReplacement(appId, backupPath);
    if (backupHr != S_OK)
    {
        return backupHr;
    }

    settings.persistence.savePermission      = Common::Settings::SettingsSavePermission::Automatic;
    settings.persistence.sourceSchemaVersion = settings.schemaVersion;
    settings.persistence.expectedFileStamp.reset();
    return SaveSettingsAndSchema(appId, settings);
}

HRESULT QueueSettingsSave(std::wstring_view appId,
                          const Common::Settings::Settings& settings,
                          std::wstring_view telemetryMetric,
                          std::wstring_view telemetryContext) noexcept
{
    try
    {
        return GetSettingsSaveCoordinator().EnqueueAsync(appId, settings, telemetryMetric, telemetryContext);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // noexcept persistence boundary: report the failed asynchronous submission and return an HRESULT.
        Debug::Error(L"SettingsHotReload: failed to queue asynchronous settings save for '{}'", appId);
        return E_FAIL;
    }
}

bool FlushQueuedSettingsSaves(DWORD timeoutMs) noexcept
{
    try
    {
        return GetSettingsSaveCoordinator().FlushFor(timeoutMs);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // noexcept shutdown/test boundary: report the unavailable flush and let the caller enforce its deadline policy.
        Debug::Error(L"SettingsHotReload: failed to flush serialized settings saves.");
        return false;
    }
}

ChangedSettingsLoadResult TryLoadChangedSettings() noexcept
{
    constexpr int kEpochRetryCount = 3;
    for (int observationAttempt = 0; observationAttempt < kEpochRetryCount; ++observationAttempt)
    {
        ChangedSettingsLoadResult result{};
        std::wstring appId;
        uint64_t sessionGeneration = 0u;
        uint64_t internalSaveEpoch = 0u;
        {
            std::scoped_lock lock(g_state.mutex);
            appId             = g_state.appId;
            sessionGeneration = g_state.sessionGeneration;
            internalSaveEpoch = g_state.internalSaveEpoch;
            if (DeferNotificationForInternalSaveLocked())
            {
                result.status = ChangedSettingsStatus::NoChange;
                result.hr     = S_OK;
                return result;
            }
        }

        if (appId.empty())
        {
            result.status = ChangedSettingsStatus::Error;
            result.hr     = HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
            return result;
        }

        Common::Settings::SettingsFileStamp stamp{};
        const HRESULT stampHr = Common::Settings::TryGetSettingsFileStamp(appId, stamp);
#ifdef ENABLE_TESTS
        const DWORD postStampDelayMs = g_debugSettingsReloadPostStampDelayMs.load(std::memory_order_acquire);
        if (postStampDelayMs != 0u)
        {
            g_debugSettingsReloadPostStampDelayActive.store(true, std::memory_order_release);
            Sleep(postStampDelayMs);
            g_debugSettingsReloadPostStampDelayActive.store(false, std::memory_order_release);
        }
#endif
        bool retryObservation = false;
        {
            std::scoped_lock lock(g_state.mutex);
            if (DeferNotificationForInternalSaveLocked())
            {
                result.status = ChangedSettingsStatus::NoChange;
                result.hr     = S_OK;
                if (stampHr == S_OK)
                {
                    result.stamp = stamp;
                }
                return result;
            }
            if (sessionGeneration != g_state.sessionGeneration)
            {
                result.status = ChangedSettingsStatus::NoChange;
                result.hr     = S_OK;
                return result;
            }
            if (internalSaveEpoch != g_state.internalSaveEpoch)
            {
                retryObservation = true;
            }
            else if (stampHr == S_OK && ShouldIgnoreStampLocked(stamp))
            {
                result.status = ChangedSettingsStatus::NoChange;
                result.hr     = S_OK;
                result.stamp  = stamp;
                return result;
            }
        }
        if (retryObservation)
        {
            continue;
        }

        if (stampHr == S_FALSE)
        {
            result.status = ChangedSettingsStatus::Missing;
            result.hr     = S_FALSE;
            return result;
        }
        if (FAILED(stampHr))
        {
            result.status = ChangedSettingsStatus::Error;
            result.hr     = stampHr;
            return result;
        }

        Common::Settings::Settings loaded;
        HRESULT loadHr = S_FALSE;
        for (int loadAttempt = 0; loadAttempt < 6; ++loadAttempt)
        {
            loadHr = Common::Settings::TryLoadSettingsNoRecovery(appId, loaded);
            if (loadHr == S_OK)
            {
                break;
            }
            if (loadHr != S_FALSE && ! IsRetryableSettingsLoadFailure(loadHr))
            {
                break;
            }
            if (loadAttempt + 1 < 6)
            {
                ::Sleep(25);
            }
        }

        retryObservation = false;
        {
            std::scoped_lock lock(g_state.mutex);
            if (DeferNotificationForInternalSaveLocked())
            {
                result.status = ChangedSettingsStatus::NoChange;
                result.hr     = S_OK;
                result.stamp  = stamp;
                return result;
            }
            if (sessionGeneration != g_state.sessionGeneration)
            {
                result.status = ChangedSettingsStatus::NoChange;
                result.hr     = S_OK;
                return result;
            }
            if (internalSaveEpoch != g_state.internalSaveEpoch)
            {
                retryObservation = true;
            }
            else if (ShouldIgnoreStampLocked(stamp))
            {
                result.status = ChangedSettingsStatus::NoChange;
                result.hr     = S_OK;
                result.stamp  = stamp;
                return result;
            }
        }
        if (retryObservation)
        {
            continue;
        }

        if (loadHr == S_OK)
        {
            result.status   = ChangedSettingsStatus::Loaded;
            result.hr       = S_OK;
            result.settings = std::move(loaded);
            result.stamp    = stamp;
            return result;
        }
        if (loadHr == S_FALSE)
        {
            result.status = ChangedSettingsStatus::Missing;
            result.hr     = S_FALSE;
            return result;
        }

        result.status = IsInvalidSettingsLoadFailure(loadHr) ? ChangedSettingsStatus::Invalid : ChangedSettingsStatus::Error;
        result.hr     = loadHr;
        result.stamp  = stamp;
        return result;
    }

    HWND retryTarget = nullptr;
    {
        std::scoped_lock lock(g_state.mutex);
        retryTarget = g_state.targetWindow;
    }
    PostSettingsFileChanged(retryTarget);
    return ChangedSettingsLoadResult{.status = ChangedSettingsStatus::NoChange, .hr = S_OK};
}

void MarkAppliedStamp(const Common::Settings::SettingsFileStamp& stamp) noexcept
{
    std::scoped_lock lock(g_state.mutex);
    g_state.lastAppliedStamp = stamp;
    g_state.lastRejectedStamp.reset();
}

void MarkRejectedStamp(const Common::Settings::SettingsFileStamp& stamp) noexcept
{
    std::scoped_lock lock(g_state.mutex);
    g_state.lastRejectedStamp = stamp;
}

void RegisterParticipant(HWND hwnd) noexcept
{
    if (! hwnd || ! IsWindow(hwnd))
    {
        return;
    }

    std::scoped_lock lock(g_state.mutex);
    g_state.participants.insert(hwnd);
}

void UnregisterParticipant(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    std::scoped_lock lock(g_state.mutex);
    g_state.participants.erase(hwnd);
}

void NotifyParticipants() noexcept
{
    std::vector<HWND> participants;
    {
        std::scoped_lock lock(g_state.mutex);
        participants.reserve(g_state.participants.size());
        for (HWND hwnd : g_state.participants)
        {
            if (hwnd && IsWindow(hwnd))
            {
                participants.push_back(hwnd);
            }
        }
    }

    for (HWND hwnd : participants)
    {
        static_cast<void>(PostMessageW(hwnd, WndMsg::kSettingsReloadedFromDisk, 0, 0));
    }
}

HRESULT PromptExternalReloadConflict(HWND targetWindow, std::wstring_view editorName, ExternalReloadChoice& outChoice) noexcept
{
    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_SETTINGS_RELOADED_FROM_DISK);
    const std::wstring message = BuildExternalReloadConflictMessage(editorName);

    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
    prompt.severity      = HOST_ALERT_WARNING;
    prompt.buttons       = HOST_PROMPT_BUTTONS_YES_NO;
    prompt.targetWindow  = targetWindow;
    prompt.title         = title.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_YES;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hrPrompt))
    {
        return hrPrompt;
    }

    outChoice = (promptResult == HOST_PROMPT_RESULT_YES) ? ExternalReloadChoice::ReloadFromDisk : ExternalReloadChoice::KeepEditing;
    return S_OK;
}

HRESULT PromptStaleSaveConflict(HWND targetWindow, std::wstring_view editorName, StaleSaveChoice& outChoice) noexcept
{
    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_SETTINGS_RELOADED_FROM_DISK);
    const std::wstring message = BuildStaleSaveConflictMessage(editorName);

    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
    prompt.severity      = HOST_ALERT_WARNING;
    prompt.buttons       = HOST_PROMPT_BUTTONS_YES_NO_CANCEL;
    prompt.targetWindow  = targetWindow;
    prompt.title         = title.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_CANCEL;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hrPrompt))
    {
        return hrPrompt;
    }

    switch (promptResult)
    {
        case HOST_PROMPT_RESULT_OK:
        case HOST_PROMPT_RESULT_YES: outChoice = StaleSaveChoice::OverwriteCurrent; break;

        case HOST_PROMPT_RESULT_NO: outChoice = StaleSaveChoice::ReloadFromDisk; break;

        case HOST_PROMPT_RESULT_CANCEL:
        case HOST_PROMPT_RESULT_NONE:
        default: outChoice = StaleSaveChoice::Cancel; break;
    }

    return S_OK;
}

void ShowInvalidReloadAlert(const std::filesystem::path& settingsPath) noexcept
{
    bool showAlert = false;
    {
        std::scoped_lock lock(g_state.mutex);
        if (! g_state.invalidAlertVisible)
        {
            g_state.invalidAlertVisible = true;
            showAlert                   = true;
        }
    }

    if (! showAlert)
    {
        return;
    }

    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_SETTINGS_RELOAD_FAILED);
    const std::wstring message = FormatStringResource(nullptr, IDS_FMT_SETTINGS_RELOAD_FAILED_KEEP_CURRENT, settingsPath.wstring());

    HostAlertRequest request{};
    request.version   = 1;
    request.sizeBytes = sizeof(request);
    request.scope     = HOST_ALERT_SCOPE_APPLICATION;
    request.modality  = HOST_ALERT_MODELESS;
    request.severity  = HOST_ALERT_WARNING;
    request.title     = title.c_str();
    request.message   = message.c_str();
    request.closable  = TRUE;

    const HRESULT hrAlert = HostShowAlert(request, InvalidAlertCookie());
    if (FAILED(hrAlert))
    {
        std::scoped_lock lock(g_state.mutex);
        g_state.invalidAlertVisible = false;
        Debug::Warning(L"SettingsHotReload: failed to show invalid settings alert (hr=0x{:08X})", static_cast<unsigned long>(hrAlert));
    }
}

void ClearInvalidReloadAlert() noexcept
{
    bool clearAlert = false;
    {
        std::scoped_lock lock(g_state.mutex);
        if (g_state.invalidAlertVisible)
        {
            g_state.invalidAlertVisible = false;
            clearAlert                  = true;
        }
    }

    if (! clearAlert)
    {
        return;
    }

    const HRESULT hrClear = HostClearAlert(HOST_ALERT_SCOPE_APPLICATION, InvalidAlertCookie());
    if (FAILED(hrClear))
    {
        Debug::Warning(L"SettingsHotReload: failed to clear invalid settings alert (hr=0x{:08X})", static_cast<unsigned long>(hrClear));
    }
}

AppTheme ResolveDialogThemeFromSettings(const Common::Settings::Settings& settings) noexcept
{
    std::wstring_view themeId = settings.theme.currentThemeId;

    const Common::Settings::ThemeDefinition* custom = nullptr;
    if (themeId.rfind(L"user/", 0) == 0)
    {
        const auto it = std::find_if(settings.theme.themes.begin(), settings.theme.themes.end(), [&](const Common::Settings::ThemeDefinition& entry) noexcept {
            return entry.id == themeId;
        });
        if (it != settings.theme.themes.end())
        {
            custom = &*it;
        }
    }

    AppTheme theme = ResolveAppThemeSelection(themeId, custom, L"RedSalamander").theme;

    ApplyUiPreferencesToTheme(settings, theme);

    return theme;
}

Common::Settings::Settings MergeDiskSettingsWithRuntimeSession(const Common::Settings::Settings& diskSettings,
                                                               const Common::Settings::Settings& runtimeSettings,
                                                               std::span<const std::wstring_view> runtimeWindowIds) noexcept
{
    Common::Settings::Settings merged = diskSettings;

    for (std::wstring_view windowId : runtimeWindowIds)
    {
        const auto runtimeIt = runtimeSettings.windows.find(std::wstring(windowId));
        if (runtimeIt != runtimeSettings.windows.end())
        {
            merged.windows[runtimeIt->first] = runtimeIt->second;
        }
    }

    if (! runtimeSettings.folders.has_value())
    {
        return merged;
    }

    if (! merged.folders.has_value())
    {
        merged.folders = Common::Settings::FoldersSettings{};
    }

    const auto& runtimeFolders = runtimeSettings.folders.value();
    auto& mergedFolders        = merged.folders.value();

    mergedFolders.active         = runtimeFolders.active;
    mergedFolders.layout         = runtimeFolders.layout;
    mergedFolders.history        = runtimeFolders.history;
    mergedFolders.historyFilters = runtimeFolders.historyFilters;

    for (const auto& runtimePane : runtimeFolders.items)
    {
        auto mergedIt = std::find_if(
            mergedFolders.items.begin(), mergedFolders.items.end(), [&](const Common::Settings::FolderPane& item) { return item.slot == runtimePane.slot; });
        if (mergedIt != mergedFolders.items.end())
        {
            mergedIt->current = runtimePane.current;
            continue;
        }

        Common::Settings::FolderPane pane;
        pane.slot    = runtimePane.slot;
        pane.current = runtimePane.current;
        mergedFolders.items.push_back(std::move(pane));
    }

    return merged;
}
} // namespace SettingsHotReload
