#pragma once
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>
#include <atomic>
#include <evntrace.h>
#include <functional>
#include <memory>
#include <string>
#include <tdh.h>
#include <thread>

#pragma comment(lib, "tdh.lib")

// Forward declaration - avoid including Helpers.h to prevent TraceLogging provider cross-DLL issues
namespace Debug
{
struct InfoParam;
}

// ETW Real-Time Listener for RedSalamanderMonitor
// Consumes TraceLogging events from the RedSalamander provider in real-time
//  to remove C4820 about padding added by the compiler to align the struct/class in memory.
class alignas(8) EtwListener
{
public:
    // Callback invoked for each debug message event
    // Parameters: InfoParam (metadata), message (payload text)
    using EventCallback = std::function<void(const Debug::InfoParam&, const std::wstring&)>;

    EtwListener();
    ~EtwListener();

    // Delete copy/move to enforce single ownership
    EtwListener(const EtwListener&)            = delete;
    EtwListener& operator=(const EtwListener&) = delete;
    EtwListener(EtwListener&&)                 = delete;
    EtwListener& operator=(EtwListener&&)      = delete;

    // Start listening for events with the given callback
    // Returns true on success, false if session couldn't start
    bool Start(EventCallback callback);

    // Stop listening and clean up resources
    void Stop();

    // Check if listener is currently running
    [[maybe_unused]] bool IsRunning() const;

#if defined(ENABLE_TESTS)
    using DebugProcessTraceFunction = ULONG (*)(void* context, TRACEHANDLE traceHandle) noexcept;
    using DebugCloseTraceFunction   = void (*)(void* context, TRACEHANDLE traceHandle) noexcept;

    void DebugStartConsumerForTesting(
        TRACEHANDLE traceHandle, DebugProcessTraceFunction processTrace, DebugCloseTraceFunction closeTrace, void* context, DWORD shutdownTimeoutMs);
#endif

    // Get last error message (if Start failed)
    std::wstring GetLastError() const
    {
        return _lastError;
    }

    // Get last Win32 error code (if Start failed)
    ULONG GetLastErrorCode() const noexcept
    {
        return _lastErrorCode;
    }

private:
    struct CallbackState;

    static constexpr wchar_t kSessionName[] = L"RedSalamanderMonitor_ETW_Session";
    static constexpr GUID kProviderGuid     = {0x440c70f6, 0x6c6b, 0x4ff7, {0x9a, 0x3f, 0x0b, 0x7d, 0xb4, 0x11, 0xb3, 0x1a}};

    // ETW callback functions
    static ULONG WINAPI BufferCallback(PEVENT_TRACE_LOGFILEW logfile);
    static VOID WINAPI EventRecordCallback(PEVENT_RECORD eventRecord);

    // Per-session callback handlers. EVENT_TRACE_LOGFILE::Context owns the routing identity.
    static void HandleEvent(CallbackState& state, PEVENT_RECORD eventRecord);

    // Extract data from TraceLogging event
    static bool ExtractEventData(PEVENT_RECORD eventRecord, Debug::InfoParam& info, std::wstring& message);

    using ProcessTraceFunction = ULONG (*)(void* context, TRACEHANDLE traceHandle) noexcept;
    static ULONG ProcessTraceConsumer(void* context, TRACEHANDLE traceHandle) noexcept;
    void StartConsumerWorker(TRACEHANDLE traceHandle,
                             const std::shared_ptr<CallbackState>& callbackState,
                             ProcessTraceFunction processTrace,
                             void* processTraceContext);

    // Member variables
    std::shared_ptr<CallbackState> _callbackState;
    TRACEHANDLE _sessionHandle;
    TRACEHANDLE _traceHandle;
    std::jthread _workerThread;
    std::wstring _lastError;
    ULONG _lastErrorCode = ERROR_SUCCESS;
#if defined(ENABLE_TESTS)
    DebugCloseTraceFunction _debugCloseTrace = nullptr;
    void* _debugTraceContext                 = nullptr;
    DWORD _shutdownTimeoutMs                 = 5'000u;
#endif

public:
    // Get buffer statistics (for diagnostics)
    struct Statistics
    {
        ULONG buffersProcessed;
        ULONG eventsProcessed;
        ULONG eventsLost;
        double eventLossRate; // Percentage of events lost
    };
    Statistics GetStatistics() const;
};
