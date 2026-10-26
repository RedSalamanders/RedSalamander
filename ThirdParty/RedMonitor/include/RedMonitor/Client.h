#pragma once

// Define REDMONITOR_DEFINE_PROVIDER in exactly one translation unit per EXE/DLL before this include.
// Own one Provider per module. Join all producers before destroying it or unloading the module.
#include <algorithm>
#include <array>
#include <atomic>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <evntrace.h>

#include <ShlObj.h>
#include <TraceLoggingProvider.h>
#include <wil/resource.h>
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "version.lib")

#pragma comment(lib, "advapi32.lib")

namespace RedMonitor
{
inline constexpr unsigned int kApiRevision   = 1;
inline constexpr size_t kMaxMessageCodeUnits = 16000;

enum class Severity : uint32_t
{
    Text    = 0,
    Error   = 1,
    Warning = 2,
    Info    = 4,
    Perf    = 8,
    Debug   = 16,
};

[[nodiscard]] constexpr uint32_t FilterBit(Severity severity) noexcept
{
    switch (severity)
    {
        case Severity::Text: return 1;
        case Severity::Error: return 2;
        case Severity::Warning: return 4;
        case Severity::Info: return 8;
        case Severity::Perf: return 16;
        case Severity::Debug: return 32;
        default: return 0;
    }
}

namespace Detail
{
[[nodiscard]] __declspec(noinline) inline ULONG WriteNormalizedEnabled(TraceLoggingHProvider provider,
                                                                       std::wstring_view application,
                                                                       Severity severity,
                                                                       std::wstring_view message,
                                                                       DWORD processId,
                                                                       DWORD threadId,
                                                                       FILETIME timestamp) noexcept
{
    std::array<wchar_t, kMaxMessageCodeUnits> storage;
    std::wstring_view payload = message;
    if (! application.empty())
    {
        storage[0] = L'[';
        std::copy(application.begin(), application.end(), storage.begin() + 1);
        storage[application.size() + 1] = L']';
        storage[application.size() + 2] = L' ';
        std::copy(message.begin(), message.end(), storage.begin() + application.size() + 3);
        payload = {storage.data(), application.size() + 3 + message.size()};
    }
    const uint64_t fileTime = (static_cast<uint64_t>(timestamp.dwHighDateTime) << 32u) | timestamp.dwLowDateTime;
    TraceLoggingWrite(provider,
                      "DebugMessage",
                      TraceLoggingLevel(TRACE_LEVEL_INFORMATION),
                      TraceLoggingKeyword(1),
                      TraceLoggingUInt32(static_cast<uint32_t>(severity), "Type"),
                      TraceLoggingUInt32(processId, "ProcessId"),
                      TraceLoggingUInt32(threadId, "ThreadId"),
                      TraceLoggingUInt64(fileTime, "FileTime"),
                      TraceLoggingCountedWideString(payload.data() ? payload.data() : L"", static_cast<USHORT>(payload.size()), "Message"),
                      TraceLoggingCountedWideString(application.data() ? application.data() : L"", static_cast<USHORT>(application.size()), "Application"));
    return ERROR_SUCCESS;
}
} // namespace Detail
// Reuse an existing module-owned provider and metadata without changing its ABI or registration lifetime.
[[nodiscard]] inline ULONG WriteNormalizedEvent(TraceLoggingHProvider provider,
                                                std::wstring_view application,
                                                Severity severity,
                                                std::wstring_view message,
                                                DWORD processId,
                                                DWORD threadId,
                                                FILETIME timestamp) noexcept
{
    if (FilterBit(severity) == 0 || application.size() > 64)
        return ERROR_INVALID_PARAMETER;
    const size_t prefix = application.empty() ? 0 : application.size() + 3;
    if (message.size() > kMaxMessageCodeUnits - prefix)
        return ERROR_BUFFER_OVERFLOW;
    if (! TraceLoggingProviderEnabled(provider, TRACE_LEVEL_INFORMATION, 1))
        return ERROR_SUCCESS;
    return Detail::WriteNormalizedEnabled(provider, application, severity, message, processId, threadId, timestamp);
}

namespace Detail
{
#if defined(REDMONITOR_DEFINE_PROVIDER)
TRACELOGGING_DEFINE_PROVIDER(g_provider, "RedSalamanderMonitor", (0x440c70f6, 0x6c6b, 0x4ff7, 0x9a, 0x3f, 0x0b, 0x7d, 0xb4, 0x11, 0xb3, 0x1a));
#else
TRACELOGGING_DECLARE_PROVIDER(g_provider);
#endif
inline std::atomic<bool> g_hasOwner{false};
} // namespace Detail

class Provider final
{
public:
    Provider() noexcept
    {
        bool expected = false;
        _owns         = Detail::g_hasOwner.compare_exchange_strong(expected, true, std::memory_order_acq_rel);
        _status       = _owns ? TraceLoggingRegister(Detail::g_provider) : ERROR_ALREADY_EXISTS;
    }

    ~Provider() noexcept
    {
        if (_owns)
        {
            if (_status == ERROR_SUCCESS)
            {
                TraceLoggingUnregister(Detail::g_provider);
            }
            Detail::g_hasOwner.store(false, std::memory_order_release);
        }
    }

    Provider(const Provider&)            = delete;
    Provider& operator=(const Provider&) = delete;
    Provider(Provider&&)                 = delete;
    Provider& operator=(Provider&&)      = delete;

    [[nodiscard]] ULONG Status() const noexcept
    {
        return _status;
    }

    [[nodiscard]] bool IsEnabled() const noexcept
    {
        return _status == ERROR_SUCCESS && TraceLoggingProviderEnabled(Detail::g_provider, TRACE_LEVEL_INFORMATION, 1);
    }

    // ERROR_SUCCESS means accepted by ETW or disabled, not guaranteed viewer delivery.
    [[nodiscard]] ULONG Write(Severity severity, std::wstring_view message) const noexcept
    {
        return WriteLabelled({}, severity, message);
    }
    [[nodiscard]] ULONG WriteLabelled(std::wstring_view application, Severity severity, std::wstring_view message) const noexcept
    {
        if (_status != ERROR_SUCCESS)
            return _status;
        if (FilterBit(severity) == 0 || application.size() > 64)
            return ERROR_INVALID_PARAMETER;
        const size_t prefix = application.empty() ? 0 : application.size() + 3;
        if (message.size() > kMaxMessageCodeUnits - prefix)
            return ERROR_BUFFER_OVERFLOW;
        if (! IsEnabled())
            return ERROR_SUCCESS;
        FILETIME timestamp{};
        GetSystemTimeAsFileTime(&timestamp);
        return WriteNormalizedEvent(Detail::g_provider, application, severity, message, GetCurrentProcessId(), GetCurrentThreadId(), timestamp);
    }

private:
    ULONG _status = ERROR_INVALID_HANDLE;
    bool _owns    = false;
};
// The label is borrowed for the module lifetime. Construct before starting producers;
// destroy only after their workers/callbacks and plugin log forwarding have stopped.
class Application final
{
public:
    explicit Application(std::wstring_view label) noexcept : _label(label)
    {
    }

    [[nodiscard]] ULONG Status() const noexcept
    {
        return _provider.Status();
    }
    [[nodiscard]] bool IsEnabled() const noexcept
    {
        return _provider.IsEnabled();
    }

    [[nodiscard]] ULONG Write(Severity severity, std::wstring_view message) const noexcept
    {
        if (FilterBit(severity) == 0 || _label.empty() || _label.size() > 64)
            return ERROR_INVALID_PARAMETER;
        if (message.size() > kMaxMessageCodeUnits - _label.size() - 3)
            return ERROR_BUFFER_OVERFLOW;
        if (! IsEnabled())
            return Status();
        return _provider.WriteLabelled(_label, severity, message);
    }

    // Bounded, counted UTF-8 from an app-owned log record. Conversion happens only with a subscriber.
    [[nodiscard]] ULONG WriteUtf8(Severity severity, std::string_view message) const noexcept
    {
        if (FilterBit(severity) == 0 || _label.empty() || _label.size() > 64)
            return ERROR_INVALID_PARAMETER;
        if (message.size() > kMaxMessageCodeUnits)
            return ERROR_BUFFER_OVERFLOW;
        if (! IsEnabled())
            return Status();
        return WriteUtf8Enabled(severity, message);
    }

private:
    [[nodiscard]] __declspec(noinline) ULONG WriteUtf8Enabled(Severity severity, std::string_view message) const noexcept
    {
        if (message.empty())
            return Write(severity, {});
        std::array<wchar_t, kMaxMessageCodeUnits> text;
        const int count =
            MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, message.data(), static_cast<int>(message.size()), text.data(), static_cast<int>(text.size()));
        if (count == 0)
            return GetLastError();
        return Write(severity, {text.data(), static_cast<size_t>(count)});
    }

    Provider _provider;
    std::wstring_view _label;
};

// Installed helper discovery/launch is startup work, never part of the logging fast path.
// Only production RedMonitor images in the packaging-owned standard locations are eligible.
[[nodiscard]] inline bool IsProductionViewer(std::wstring_view executable)
{
    if (executable.empty() || executable.size() >= 32767)
        return false;
    const std::wstring path(executable);
    DWORD ignored{};
    const DWORD size = GetFileVersionInfoSizeW(path.c_str(), &ignored);
    if (size == 0 || size > 1024 * 1024)
        return false;
    std::vector<std::byte> data(size);
    if (! GetFileVersionInfoW(path.c_str(), 0, size, data.data()))
        return false;
    VS_FIXEDFILEINFO* fixed{};
    UINT count{};
    if (! VerQueryValueW(data.data(), L"\\", reinterpret_cast<void**>(&fixed), &count) || count < sizeof(*fixed) ||
        (fixed->dwFileFlags & (VS_FF_DEBUG | VS_FF_SPECIALBUILD)) != 0)
        return false;
    wchar_t* product{};
    if (! VerQueryValueW(data.data(), L"\\StringFileInfo\\040904b0\\ProductName", reinterpret_cast<void**>(&product), &count) || count == 0 ||
        std::wstring_view(product, count - 1) != L"RedMonitor")
        return false;
    wchar_t* special{};
    if (VerQueryValueW(data.data(), L"\\StringFileInfo\\040904b0\\SpecialBuild", reinterpret_cast<void**>(&special), &count) && count > 1)
        return false;
    return true;
}
[[nodiscard]] inline std::wstring FindInstalledViewer()
{
    constexpr std::array folders{&FOLDERID_LocalAppData, &FOLDERID_ProgramFiles};
    for (size_t index = 0; index < folders.size(); ++index)
    {
        wil::unique_cotaskmem_string folder;
        if (FAILED(SHGetKnownFolderPath(*folders[index], KF_FLAG_DEFAULT, nullptr, folder.put())))
            continue;
        std::wstring path(folder.get());
        path += index == 0 ? L"\\Programs\\RedMonitor\\RedMonitor.exe" : L"\\RedMonitor\\RedMonitor.exe";
        if (IsProductionViewer(path))
            return path;
    }
    return {};
}
[[nodiscard]] constexpr bool IsViewerCommand(std::wstring_view command) noexcept
{
    const auto first = command.find_first_not_of(L" \t\r\n");
    if (first == std::wstring_view::npos)
        return false;
    const auto last = command.find_last_not_of(L" \t\r\n");
    return command.substr(first, last - first + 1) == L"--redmonitor";
}
[[nodiscard]] constexpr bool IsOrdinaryStartup(std::wstring_view command) noexcept
{
    return command.find_first_not_of(L" \t\r\n") == std::wstring_view::npos;
}
[[nodiscard]] inline DWORD LaunchInstalledViewer()
{
    const auto path = FindInstalledViewer();
    if (path.empty())
        return ERROR_FILE_NOT_FOUND;
    // Preserve the viewer's independent lifetime and single-instance policy.
    wil::unique_handle existing(OpenMutexW(SYNCHRONIZE, FALSE, L"Local\\RedSalamanderMonitor_Instance"));
    if (existing)
        return ERROR_SUCCESS;
    auto command = L"\"" + path + L"\"";
    STARTUPINFOW startup{};
    startup.cb = sizeof(startup);
    PROCESS_INFORMATION process{};
    if (! CreateProcessW(path.c_str(), command.data(), nullptr, nullptr, FALSE, 0, nullptr, nullptr, &startup, &process))
        return GetLastError();
    wil::unique_handle thread(process.hThread);
    wil::unique_handle owner(process.hProcess);
    return ERROR_SUCCESS;
}

} // namespace RedMonitor
