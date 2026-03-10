#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <shellapi.h>

#include <cstdint>
#include <string>
#include <string_view>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"
#include "SearchServiceBroker.h"

#pragma comment(lib, "Advapi32.lib")

namespace
{
struct ParsedArguments final
{
    bool foregroundMode                = false;
    uint32_t protocolVersion           = SearchServiceBroker::kProtocolVersion;
    uint32_t maxRequestsBeforeExit     = 0u;
    uint32_t disconnectAfterBatches    = 0u;
    std::wstring pipeName;
};

SERVICE_STATUS_HANDLE g_serviceStatusHandle = nullptr;
SERVICE_STATUS g_serviceStatus{};
wil::unique_event g_serviceStopEvent;

void UpdateServiceStatus(DWORD currentState, DWORD win32ExitCode, DWORD waitHint) noexcept
{
    if (! g_serviceStatusHandle)
    {
        return;
    }

    g_serviceStatus.dwServiceType             = SERVICE_WIN32_OWN_PROCESS;
    g_serviceStatus.dwCurrentState            = currentState;
    g_serviceStatus.dwWin32ExitCode           = win32ExitCode;
    g_serviceStatus.dwWaitHint                = waitHint;
    g_serviceStatus.dwControlsAccepted        = currentState == SERVICE_START_PENDING ? 0u : SERVICE_ACCEPT_STOP | SERVICE_ACCEPT_SHUTDOWN;
    g_serviceStatus.dwCheckPoint              = (currentState == SERVICE_RUNNING || currentState == SERVICE_STOPPED) ? 0u : 1u;
    static_cast<void>(::SetServiceStatus(g_serviceStatusHandle, &g_serviceStatus));
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
        if (arg == L"--run-foreground")
        {
            parsed.foregroundMode = true;
            continue;
        }

        if (arg.rfind(L"--pipe-name=", 0) == 0)
        {
            parsed.pipeName = std::wstring(arg.substr(std::wstring_view(L"--pipe-name=").size()));
            continue;
        }

        if (arg.rfind(L"--protocol-version=", 0) == 0)
        {
            uint32_t value = 0u;
            if (TryParseUnsigned(arg.substr(std::wstring_view(L"--protocol-version=").size()), value))
            {
                parsed.protocolVersion = value;
            }
            continue;
        }

        if (arg.rfind(L"--max-requests=", 0) == 0)
        {
            uint32_t value = 0u;
            if (TryParseUnsigned(arg.substr(std::wstring_view(L"--max-requests=").size()), value))
            {
                parsed.maxRequestsBeforeExit = value;
            }
            continue;
        }

        if (arg.rfind(L"--disconnect-after-batches=", 0) == 0)
        {
            uint32_t value = 0u;
            if (TryParseUnsigned(arg.substr(std::wstring_view(L"--disconnect-after-batches=").size()), value))
            {
                parsed.disconnectAfterBatches = value;
            }
            continue;
        }
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

    g_serviceStopEvent.create();
    UpdateServiceStatus(SERVICE_START_PENDING, NO_ERROR, 5000u);

    SearchServiceBroker::ServerOptions options{};
    const HRESULT hr = SearchServiceBroker::RunServer(options, g_serviceStopEvent.get(), nullptr);
    UpdateServiceStatus(SERVICE_STOPPED, FAILED(hr) ? static_cast<DWORD>(HRESULT_CODE(hr)) : NO_ERROR, 0u);
}
} // namespace

int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR, int)
{
    const ParsedArguments parsed = ParseArguments();
    if (parsed.foregroundMode)
    {
        SearchServiceBroker::ServerOptions options{};
        options.pipeName               = parsed.pipeName;
        options.protocolVersion        = parsed.protocolVersion;
        options.maxRequestsBeforeExit  = parsed.maxRequestsBeforeExit;
        options.disconnectAfterBatches = parsed.disconnectAfterBatches;
        const HRESULT hr = SearchServiceBroker::RunServer(options, nullptr, nullptr);
        return FAILED(hr) ? 1 : 0;
    }

    SERVICE_TABLE_ENTRYW table[] = {
        {const_cast<LPWSTR>(SearchServiceBroker::kServiceName), &ServiceMain},
        {nullptr, nullptr},
    };

    if (::StartServiceCtrlDispatcherW(table) == 0)
    {
        return 1;
    }

    return 0;
}
