#include "EtwListener.h"
#include "Helpers.h" // For Debug::InfoParam definition
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

// Static instance pointer for callback routing
EtwListener* EtwListener::s_instance = nullptr;

EtwListener::EtwListener() : _sessionHandle(INVALID_PROCESSTRACE_HANDLE), _traceHandle(INVALID_PROCESSTRACE_HANDLE), _isRunning(false)
{
}

EtwListener::~EtwListener()
{
    Stop();
}

bool EtwListener::Start(EventCallback callback)
{
    _lastErrorCode = ERROR_SUCCESS;
    _lastError.clear();

    if (_isRunning.load())
    {
        _lastError     = L"Listener is already running";
        _lastErrorCode = ERROR_ALREADY_EXISTS;
        return false;
    }

    if (! callback)
    {
        _lastError     = L"Callback function is null";
        _lastErrorCode = ERROR_INVALID_PARAMETER;
        return false;
    }

    _userCallback = std::move(callback);
    s_instance    = this;

    // Stop any existing session with the same name
    std::vector<BYTE> sessionBuffer(sizeof(EVENT_TRACE_PROPERTIES) + (wcslen(kSessionName) + 1) * sizeof(wchar_t));
    auto* props                    = reinterpret_cast<PEVENT_TRACE_PROPERTIES>(sessionBuffer.data());
    props->Wnode.BufferSize        = static_cast<ULONG>(sessionBuffer.size());
    props->LoggerNameOffset        = sizeof(EVENT_TRACE_PROPERTIES);
    const ULONG stopExistingResult = ControlTrace(0, kSessionName, props, EVENT_TRACE_CONTROL_STOP);

    // Create new session properties with optimized buffer configuration
    // Buffer settings tuned for high-frequency event generation:
    //   MinimumBuffers: 8 (vs. default 4) for better baseline capacity
    //   MaximumBuffers: 128 (vs. 64) to handle burst traffic
    //   BufferSize: 256 KB (vs. default 64 KB) for reduced overhead
    //   Total capacity: 8-32 MB (vs. previous 256 KB-4 MB)
    ZeroMemory(props, sessionBuffer.size());
    props->Wnode.BufferSize    = static_cast<ULONG>(sessionBuffer.size());
    props->Wnode.Flags         = WNODE_FLAG_TRACED_GUID;
    props->Wnode.ClientContext = 1;                                       // Use QPC for timestamp
    props->Wnode.Guid          = GUID{0, 0, 0, {0, 0, 0, 0, 0, 0, 0, 0}}; // GUID_NULL equivalent
    props->LogFileMode         = EVENT_TRACE_REAL_TIME_MODE;
    props->LoggerNameOffset    = sizeof(EVENT_TRACE_PROPERTIES);
    props->BufferSize          = 256; // 256 KB per buffer (increased from default 64 KB)
    props->MinimumBuffers      = 8;   // Increased from 4
    props->MaximumBuffers      = 128; // Increased from 64
    props->FlushTimer          = 1;   // Flush every second
    wcscpy_s(reinterpret_cast<wchar_t*>(sessionBuffer.data() + props->LoggerNameOffset), wcslen(kSessionName) + 1, kSessionName);

    // Start the trace session
    ULONG result = StartTrace(&_sessionHandle, kSessionName, props);
    if (result != ERROR_SUCCESS)
    {
        _lastErrorCode = result;

        if (result == ERROR_ALREADY_EXISTS)
        {
            if (stopExistingResult == ERROR_ACCESS_DENIED)
            {
                _lastErrorCode = ERROR_ACCESS_DENIED;
                _lastError     = L"Existing ETW session could not be stopped (access denied)";
            }
            else
            {
                _lastError = L"ETW session already exists (another instance may be running)";
            }
        }
        else if (result == ERROR_ACCESS_DENIED)
        {
            _lastError = L"ETW session requires administrator privileges or proper ACLs";
        }
        else
        {
            _lastError = std::format(L"Failed to start ETW session: error 0x{:08X}", static_cast<unsigned long>(result));
        }
        _sessionHandle = INVALID_PROCESSTRACE_HANDLE;
        return false;
    }

    // Enable the provider
    result = EnableTraceEx2(_sessionHandle,
                            &kProviderGuid,
                            EVENT_CONTROL_CODE_ENABLE_PROVIDER,
                            TRACE_LEVEL_VERBOSE,
                            0xFFFFFFFFFFFFFFFF, // Match any keyword
                            0,                  // Match all keywords
                            0,                  // Timeout (use default)
                            nullptr);

    if (result != ERROR_SUCCESS)
    {
        _lastErrorCode = result;
        _lastError     = std::format(L"Failed to enable ETW provider: error 0x{:08X}", static_cast<unsigned long>(result));
        ControlTrace(_sessionHandle, nullptr, props, EVENT_TRACE_CONTROL_STOP);
        _sessionHandle = INVALID_PROCESSTRACE_HANDLE;
        return false;
    }

    // Open the trace for processing
    EVENT_TRACE_LOGFILE logfile{};
    logfile.LoggerName          = const_cast<wchar_t*>(kSessionName);
    logfile.ProcessTraceMode    = PROCESS_TRACE_MODE_REAL_TIME | PROCESS_TRACE_MODE_EVENT_RECORD;
    logfile.EventRecordCallback = EventRecordCallback;
    logfile.BufferCallback      = BufferCallback;

    _traceHandle = OpenTrace(&logfile);
    if (_traceHandle == INVALID_PROCESSTRACE_HANDLE)
    {
        const ULONG errorCode = ::GetLastError();
        _lastErrorCode        = errorCode;
        _lastError            = std::format(L"Failed to open ETW trace: error 0x{:08X}", static_cast<unsigned long>(errorCode));
        ControlTrace(_sessionHandle, nullptr, props, EVENT_TRACE_CONTROL_STOP);
        _sessionHandle = INVALID_PROCESSTRACE_HANDLE;
        return false;
    }

    // Start worker thread to process events
    _isRunning.store(true);
    _workerThread = std::jthread([this]([[maybe_unused]] std::stop_token st) { ProcessTraceThread(); });

    return true;
}

void EtwListener::Stop()
{
    const bool wasRunning = _isRunning.exchange(false);

    if (s_instance == this)
    {
        s_instance = nullptr;
    }

    // Close the trace handle (this will cause ProcessTrace to return)
    if (_traceHandle != INVALID_PROCESSTRACE_HANDLE)
    {
        CloseTrace(_traceHandle);
        _traceHandle = INVALID_PROCESSTRACE_HANDLE;
    }

    if (wasRunning)
    {
        // Wait for worker thread to finish (jthread destructor handles this)
        _workerThread = std::jthread();
    }

    // Stop the trace session
    if (_sessionHandle != INVALID_PROCESSTRACE_HANDLE)
    {
        std::vector<BYTE> sessionBuffer(sizeof(EVENT_TRACE_PROPERTIES) + (wcslen(kSessionName) + 1) * sizeof(wchar_t));
        auto* props             = reinterpret_cast<PEVENT_TRACE_PROPERTIES>(sessionBuffer.data());
        props->Wnode.BufferSize = static_cast<ULONG>(sessionBuffer.size());
        props->LoggerNameOffset = sizeof(EVENT_TRACE_PROPERTIES);
        ControlTrace(_sessionHandle, kSessionName, props, EVENT_TRACE_CONTROL_STOP);
        _sessionHandle = INVALID_PROCESSTRACE_HANDLE;
    }
}

void EtwListener::ProcessTraceThread()
{
    // ProcessTrace blocks until CloseTrace is called or an error occurs
    const ULONG result = ProcessTrace(&_traceHandle, 1, nullptr, nullptr);

    if (result != ERROR_SUCCESS && result != ERROR_CANCELLED)
    {
#ifdef _DEBUG
        auto msg = std::format(L"ProcessTrace ended with error: 0x{:08X}\n", static_cast<unsigned long>(result));
        OutputDebugStringW(msg.c_str());
#endif
    }

    _isRunning.store(false);
}

ULONG WINAPI EtwListener::BufferCallback(PEVENT_TRACE_LOGFILEW logfile)
{
    // Track buffer statistics
    if (s_instance && logfile)
    {
        s_instance->_buffersProcessed.fetch_add(1u, std::memory_order_relaxed);
        s_instance->_eventsLost.fetch_add(logfile->EventsLost, std::memory_order_relaxed);
    }

    // Return TRUE to continue processing
    return TRUE;
}

VOID WINAPI EtwListener::EventRecordCallback(PEVENT_RECORD eventRecord)
{
    if (s_instance)
    {
        s_instance->HandleEvent(eventRecord);
    }
}

void EtwListener::HandleEvent(PEVENT_RECORD eventRecord)
{
    if (! eventRecord || ! _userCallback)
    {
        return;
    }

    // Verify this is from our provider
    if (memcmp(&eventRecord->EventHeader.ProviderId, &kProviderGuid, sizeof(GUID)) != 0)
    {
        return;
    }

    // Track events processed
    _eventsProcessed.fetch_add(1u, std::memory_order_relaxed);

    // Extract event data
    Debug::InfoParam info{};
    std::wstring message;

    if (ExtractEventData(eventRecord, info, message))
    {
        _userCallback(info, message);
    }
}

bool EtwListener::ExtractEventData(PEVENT_RECORD eventRecord, Debug::InfoParam& info, std::wstring& message)
{
    // Get event information using TDH (Trace Data Helper)
    DWORD bufferSize = 0;
    ULONG result     = TdhGetEventInformation(eventRecord, 0, nullptr, nullptr, &bufferSize);

    if (result != ERROR_INSUFFICIENT_BUFFER)
    {
        return false;
    }

    std::vector<BYTE> buffer(bufferSize);
    auto* eventInfo = reinterpret_cast<PTRACE_EVENT_INFO>(buffer.data());

    result = TdhGetEventInformation(eventRecord, 0, nullptr, eventInfo, &bufferSize);
    if (result != ERROR_SUCCESS)
    {
        return false;
    }

    // Convert timestamp to FILETIME
    FILETIME ft{};
    ft.dwLowDateTime  = eventRecord->EventHeader.TimeStamp.LowPart;
    ft.dwHighDateTime = static_cast<DWORD>(eventRecord->EventHeader.TimeStamp.HighPart);

    // Initialize InfoParam
    info.time      = ft;
    info.processID = eventRecord->EventHeader.ProcessId;
    info.threadID  = eventRecord->EventHeader.ThreadId;
    info.type      = Debug::InfoParam::Type::Info; // Default, will be overwritten

    std::wstring perfScopeName;
    std::wstring perfScopeDetail;
    uint64_t perfDurationUs = 0;
    uint64_t perfValue0     = 0;
    uint64_t perfValue1     = 0;
    uint32_t perfHr         = 0;

    // Extract properties
    for (ULONG i = 0; i < eventInfo->TopLevelPropertyCount; ++i)
    {
        const EVENT_PROPERTY_INFO& propertyInfo = eventInfo->EventPropertyInfoArray[i];
        const wchar_t* propertyName             = reinterpret_cast<const wchar_t*>(reinterpret_cast<BYTE*>(eventInfo) + propertyInfo.NameOffset);

        PROPERTY_DATA_DESCRIPTOR dataDescriptor{};
        dataDescriptor.PropertyName = reinterpret_cast<ULONGLONG>(propertyName);
        dataDescriptor.ArrayIndex   = ULONG_MAX;

        DWORD propertySize = 0;
        result             = TdhGetPropertySize(eventRecord, 0, nullptr, 1, &dataDescriptor, &propertySize);
        if (result != ERROR_SUCCESS)
        {
            continue;
        }

        std::vector<BYTE> propertyBuffer(propertySize);
        result = TdhGetProperty(eventRecord, 0, nullptr, 1, &dataDescriptor, propertySize, propertyBuffer.data());
        if (result != ERROR_SUCCESS)
        {
            continue;
        }

        // Match property names from TraceLoggingWrite call
        if (_wcsicmp(propertyName, L"Type") == 0 && propertySize == sizeof(UINT32))
        {
            const UINT32 typeValue = *reinterpret_cast<UINT32*>(propertyBuffer.data());
            info.type              = static_cast<Debug::InfoParam::Type>(typeValue);
        }
        else if (_wcsicmp(propertyName, L"Name") == 0 || _wcsicmp(propertyName, L"Detail") == 0 || _wcsicmp(propertyName, L"Message") == 0)
        {
            // Counted wide string from TraceLoggingCountedWideString.
            std::wstring extracted;
            if (propertySize >= sizeof(USHORT))
            {
                const USHORT lengthInBytes = *reinterpret_cast<const USHORT*>(propertyBuffer.data());
                if (lengthInBytes > 0 && lengthInBytes <= propertySize - sizeof(USHORT))
                {
                    const wchar_t* strPtr  = reinterpret_cast<const wchar_t*>(propertyBuffer.data() + sizeof(USHORT));
                    const size_t charCount = static_cast<size_t>(lengthInBytes) / sizeof(wchar_t);
                    extracted.assign(strPtr, charCount);

                    // Trim at first NUL if present.
                    const size_t nulPos = extracted.find(L'\0');
                    if (nulPos != std::wstring::npos)
                    {
                        extracted.resize(nulPos);
                    }
                }
            }

            if (! extracted.empty())
            {
                if (_wcsicmp(propertyName, L"Message") == 0)
                {
                    message = std::move(extracted);
                }
                else if (_wcsicmp(propertyName, L"Name") == 0)
                {
                    perfScopeName = std::move(extracted);
                }
                else
                {
                    perfScopeDetail = std::move(extracted);
                }
            }
        }
        else if (_wcsicmp(propertyName, L"ProcessId") == 0 && propertySize == sizeof(UINT32))
        {
            info.processID = *reinterpret_cast<UINT32*>(propertyBuffer.data());
        }
        else if (_wcsicmp(propertyName, L"ThreadId") == 0 && propertySize == sizeof(UINT32))
        {
            info.threadID = *reinterpret_cast<UINT32*>(propertyBuffer.data());
        }
        else if (_wcsicmp(propertyName, L"FileTime") == 0 && propertySize == sizeof(UINT64))
        {
            const UINT64 fileTime    = *reinterpret_cast<UINT64*>(propertyBuffer.data());
            info.time.dwLowDateTime  = static_cast<DWORD>(fileTime & 0xFFFFFFFF);
            info.time.dwHighDateTime = static_cast<DWORD>(fileTime >> 32);
        }
        else if (_wcsicmp(propertyName, L"DurationUs") == 0 && propertySize == sizeof(UINT64))
        {
            perfDurationUs = *reinterpret_cast<const UINT64*>(propertyBuffer.data());
        }
        else if (_wcsicmp(propertyName, L"Value0") == 0 && propertySize == sizeof(UINT64))
        {
            perfValue0 = *reinterpret_cast<const UINT64*>(propertyBuffer.data());
        }
        else if (_wcsicmp(propertyName, L"Value1") == 0 && propertySize == sizeof(UINT64))
        {
            perfValue1 = *reinterpret_cast<const UINT64*>(propertyBuffer.data());
        }
        else if (_wcsicmp(propertyName, L"Hr") == 0 && propertySize == sizeof(UINT32))
        {
            perfHr = *reinterpret_cast<const UINT32*>(propertyBuffer.data());
        }
    }

    if (message.empty() && ! perfScopeName.empty())
    {
        info.type = Debug::InfoParam::Type::Debug;

        constexpr UINT64 kPerfWarningUs = 500'000;   // 500ms
        constexpr UINT64 kPerfErrorUs   = 1'000'000; // 1s

        std::wstring_view perfEmoji;
        if (perfDurationUs >= kPerfErrorUs)
        {
            perfEmoji = L"❌ ";
        }
        else if (perfDurationUs >= kPerfWarningUs)
        {
            perfEmoji = L"⚠️ ";
        }
        else
        {
            perfEmoji = L"";
        }

        const UINT64 perfDurationMs          = perfDurationUs / 1000;
        const UINT64 perfDurationUsRemainder = perfDurationUs % 1000;
        const std::wstring perfDurationText  = std::format(L"{}.{:03}ms", perfDurationMs, perfDurationUsRemainder);
        if (! perfScopeDetail.empty())
        {
            message = std::format(
                L"[perf] {}{} ({}) {} v0={} v1={} hr=0x{:08X}", perfEmoji, perfScopeName, perfScopeDetail, perfDurationText, perfValue0, perfValue1, perfHr);
        }
        else
        {
            message = std::format(L"[perf] {}{} {} v0={} v1={} hr=0x{:08X}", perfEmoji, perfScopeName, perfDurationText, perfValue0, perfValue1, perfHr);
        }
    }

    return ! message.empty();
}

EtwListener::Statistics EtwListener::GetStatistics() const
{
    const ULONG eventsProcessed = _eventsProcessed.load(std::memory_order_relaxed);
    const ULONG eventsLost      = _eventsLost.load(std::memory_order_relaxed);
    const ULONG totalEvents     = eventsProcessed + eventsLost;

    Statistics stats{};
    stats.buffersProcessed = _buffersProcessed.load(std::memory_order_relaxed);
    stats.eventsProcessed  = eventsProcessed;
    stats.eventsLost       = eventsLost;
    stats.eventLossRate    = totalEvents > 0 ? (static_cast<double>(eventsLost) / static_cast<double>(totalEvents)) * 100.0 : 0.0;

    return stats;
}
