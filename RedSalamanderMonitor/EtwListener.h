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
    [[maybe_unused]] bool IsRunning() const
    {
        return _isRunning;
    }

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
    static constexpr wchar_t kSessionName[] = L"RedSalamanderMonitor_ETW_Session";
    static constexpr GUID kProviderGuid     = {0x440c70f6, 0x6c6b, 0x4ff7, {0x9a, 0x3f, 0x0b, 0x7d, 0xb4, 0x11, 0xb3, 0x1a}};

    // ETW callback functions
    static ULONG WINAPI BufferCallback(PEVENT_TRACE_LOGFILEW logfile);
    static VOID WINAPI EventRecordCallback(PEVENT_RECORD eventRecord);

    // Instance method to handle events
    void HandleEvent(PEVENT_RECORD eventRecord);

    // Extract data from TraceLogging event
    bool ExtractEventData(PEVENT_RECORD eventRecord, Debug::InfoParam& info, std::wstring& message);

    // Worker thread function
    void ProcessTraceThread();

    // Member variables
    EventCallback _userCallback;
    TRACEHANDLE _sessionHandle;
    TRACEHANDLE _traceHandle;
    std::jthread _workerThread;
    std::atomic<bool> _isRunning;
    std::wstring _lastError;
    ULONG _lastErrorCode = ERROR_SUCCESS;

    // Buffer statistics for monitoring
    std::atomic<ULONG> _buffersProcessed{0};
    std::atomic<ULONG> _eventsProcessed{0};
    std::atomic<ULONG> _eventsLost{0};

    // Static instance pointer for callbacks (atomic: written on UI thread, read from ETW worker thread)
    static std::atomic<EtwListener*> s_instance;

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
