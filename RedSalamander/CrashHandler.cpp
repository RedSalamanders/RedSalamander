#include "CrashHandler.h"

#include "Framework.h"

#include <atomic>
#include <chrono>
#include <cstdint>
#include <exception>
#include <filesystem>
#include <string>

#include <dbghelp.h>
#include <shellapi.h>
#include <shlobj_core.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"
#include "HostServices.h"
#include "resource.h"

#pragma comment(lib, "Dbghelp.lib")

namespace CrashHandler
{
namespace
{
constexpr wchar_t kCompanyDirName[] = L"RedSalamander";
constexpr wchar_t kCrashDirName[]   = L"Crashes";
constexpr wchar_t kMarkerFileName[] = L"last_crash.txt";

std::atomic<bool> g_installed{false};

[[nodiscard]] std::filesystem::path GetLocalAppDataPath() noexcept
{
    wil::unique_cotaskmem_string localAppData;
    const HRESULT hr = SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, localAppData.put());
    if (SUCCEEDED(hr) && localAppData)
    {
        return std::filesystem::path(localAppData.get());
    }

    const DWORD required = GetEnvironmentVariableW(L"LOCALAPPDATA", nullptr, 0);
    if (required == 0)
    {
        return {};
    }

    std::wstring buffer(static_cast<size_t>(required), L'\0');
    const DWORD written = GetEnvironmentVariableW(L"LOCALAPPDATA", buffer.data(), required);
    if (written == 0 || written >= required)
    {
        return {};
    }

    buffer.resize(static_cast<size_t>(written));
    return std::filesystem::path(buffer);
}

[[nodiscard]] std::filesystem::path GetCrashDirectory() noexcept
{
    const std::filesystem::path base = GetLocalAppDataPath();
    if (base.empty())
    {
        return {};
    }

    return base / kCompanyDirName / kCrashDirName;
}

[[nodiscard]] std::filesystem::path GetCrashMarkerPath() noexcept
{
    const std::filesystem::path dir = GetCrashDirectory();
    if (dir.empty())
    {
        return {};
    }
    return dir / kMarkerFileName;
}

[[nodiscard]] HRESULT EnsureDirectoryExists(const std::filesystem::path& dir) noexcept
{
    if (dir.empty())
    {
        return E_INVALIDARG;
    }

    std::error_code ec;
    if (std::filesystem::exists(dir, ec))
    {
        return S_OK;
    }
    std::filesystem::create_directories(dir, ec);
    return ec ? HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value())) : S_OK;
}

[[nodiscard]] std::filesystem::path BuildDumpPath(const std::filesystem::path& dir) noexcept
{
    SYSTEMTIME st{};
    GetLocalTime(&st);

    const DWORD pid = GetCurrentProcessId();

    wchar_t fileName[128]{};
    constexpr size_t fileNameMax = (sizeof(fileName) / sizeof(fileName[0])) - 1;
    const auto r                 = std::format_to_n(fileName,
                                    fileNameMax,
                                    L"RedSalamander-{:04}{:02}{:02}-{:02}{:02}{:02}-p{}.dmp",
                                    static_cast<unsigned>(st.wYear),
                                    static_cast<unsigned>(st.wMonth),
                                    static_cast<unsigned>(st.wDay),
                                    static_cast<unsigned>(st.wHour),
                                    static_cast<unsigned>(st.wMinute),
                                    static_cast<unsigned>(st.wSecond),
                                    static_cast<unsigned long>(pid));
    const size_t written         = (r.size < fileNameMax) ? r.size : fileNameMax;
    fileName[written]            = L'\0';

    return dir / fileName;
}

[[nodiscard]] HRESULT WriteMarkerFile(const std::filesystem::path& markerPath, std::wstring_view dumpPath) noexcept
{
    if (markerPath.empty())
    {
        return E_INVALIDARG;
    }

    wil::unique_handle file(CreateFileW(markerPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const wchar_t bom = 0xFEFF;
    DWORD written     = 0;
    if (! WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const DWORD bytes = static_cast<DWORD>(dumpPath.size() * sizeof(wchar_t));
    if (bytes > 0)
    {
        if (! WriteFile(file.get(), dumpPath.data(), bytes, &written, nullptr))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }

    static_cast<void>(FlushFileBuffers(file.get()));
    return S_OK;
}

[[nodiscard]] HRESULT WriteMiniDumpFile(const std::filesystem::path& dumpPath, EXCEPTION_POINTERS* exceptionPointers) noexcept
{
    if (dumpPath.empty())
    {
        return E_INVALIDARG;
    }

    wil::unique_handle file(CreateFileW(dumpPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    MINIDUMP_EXCEPTION_INFORMATION mei{};
    mei.ThreadId          = GetCurrentThreadId();
    mei.ExceptionPointers = exceptionPointers;
    mei.ClientPointers    = FALSE;

    constexpr MINIDUMP_TYPE kDumpType = static_cast<MINIDUMP_TYPE>(MiniDumpWithIndirectlyReferencedMemory | MiniDumpScanMemory | MiniDumpWithThreadInfo |
                                                                   MiniDumpWithUnloadedModules | MiniDumpWithHandleData);

    const BOOL ok = MiniDumpWriteDump(GetCurrentProcess(), GetCurrentProcessId(), file.get(), kDumpType, exceptionPointers ? &mei : nullptr, nullptr, nullptr);
    if (! ok)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    static_cast<void>(FlushFileBuffers(file.get()));
    return S_OK;
}

[[nodiscard]] std::wstring ToWide(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring out(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), out.data(), required);
    if (written <= 0)
    {
        return {};
    }

    return out;
}

[[nodiscard]] std::wstring BuildStackTraceText(EXCEPTION_POINTERS* exceptionPointers) noexcept
{
    // Mandatory: crash path `noexcept` boundary. Stack trace generation is best-effort.
    try
    {
        HANDLE process = GetCurrentProcess();

        std::filesystem::path exeDir;
        {
            wchar_t exePath[MAX_PATH]{};
            const DWORD len = GetModuleFileNameW(nullptr, exePath, static_cast<DWORD>(std::size(exePath)));
            if (len > 0 && len < std::size(exePath))
            {
                exeDir = std::filesystem::path(exePath).parent_path();
            }
        }

        SymSetOptions(SYMOPT_DEFERRED_LOADS | SYMOPT_UNDNAME | SYMOPT_LOAD_LINES);
        if (! SymInitialize(process, nullptr, TRUE))
        {
            // Still return a minimal report even without symbols.
        }
        else if (! exeDir.empty())
        {
            const std::wstring searchPath = exeDir.wstring();
            static_cast<void>(SymSetSearchPathW(process, searchPath.c_str()));
        }

        CONTEXT context{};
        if (exceptionPointers && exceptionPointers->ContextRecord)
        {
            context = *exceptionPointers->ContextRecord;
        }
        else
        {
            RtlCaptureContext(&context);
        }

        DWORD exceptionCode    = 0;
        void* exceptionAddress = nullptr;
        if (exceptionPointers && exceptionPointers->ExceptionRecord)
        {
            exceptionCode    = exceptionPointers->ExceptionRecord->ExceptionCode;
            exceptionAddress = exceptionPointers->ExceptionRecord->ExceptionAddress;
        }

        std::wstring out;
        out.reserve(16 * 1024);
        out += std::format(L"ExceptionCode=0x{:08X}\r\n", exceptionCode);
        out += std::format(L"ExceptionAddress={}\r\n", exceptionAddress);
        out += std::format(L"ProcessId={}\r\n", GetCurrentProcessId());
        out += std::format(L"ThreadId={}\r\n", GetCurrentThreadId());

        out += L"\r\nCallstack:\r\n";

        STACKFRAME64 frame{};

#if defined(_M_X64)
        DWORD machine          = IMAGE_FILE_MACHINE_AMD64;
        frame.AddrPC.Offset    = context.Rip;
        frame.AddrFrame.Offset = context.Rbp;
        frame.AddrStack.Offset = context.Rsp;
#elif defined(_M_IX86)
        DWORD machine          = IMAGE_FILE_MACHINE_I386;
        frame.AddrPC.Offset    = context.Eip;
        frame.AddrFrame.Offset = context.Ebp;
        frame.AddrStack.Offset = context.Esp;
#else
        DWORD machine = 0;
#endif
        frame.AddrPC.Mode    = AddrModeFlat;
        frame.AddrFrame.Mode = AddrModeFlat;
        frame.AddrStack.Mode = AddrModeFlat;

        HANDLE thread = GetCurrentThread();

        for (unsigned int i = 0; i < 64u; ++i)
        {
            if (machine == 0)
            {
                break;
            }

            const BOOL ok = StackWalk64(machine, process, thread, &frame, &context, nullptr, SymFunctionTableAccess64, SymGetModuleBase64, nullptr);
            if (! ok || frame.AddrPC.Offset == 0)
            {
                break;
            }

            const DWORD64 addr = frame.AddrPC.Offset;

            char symbolBuffer[sizeof(SYMBOL_INFO) + MAX_SYM_NAME]{};
            auto* symbol         = reinterpret_cast<SYMBOL_INFO*>(symbolBuffer);
            symbol->SizeOfStruct = sizeof(SYMBOL_INFO);
            symbol->MaxNameLen   = MAX_SYM_NAME;

            std::wstring name    = L"(unknown)";
            DWORD64 displacement = 0;
            if (SymFromAddr(process, addr, &displacement, symbol))
            {
                name = ToWide(symbol->Name);
                if (name.empty())
                {
                    name = L"(unknown)";
                }
            }

            IMAGEHLP_LINE64 line{};
            line.SizeOfStruct = sizeof(line);
            DWORD lineDisp    = 0;
            std::wstring fileLine;
            if (SymGetLineFromAddr64(process, addr, &lineDisp, &line) && line.FileName)
            {
                const std::wstring file = ToWide(line.FileName);
                fileLine                = std::format(L" {}:{}(+{})", file, line.LineNumber, lineDisp);
            }

            out += std::format(L"{:02} 0x{:016X} {}+0x{:X}{}\r\n", i, addr, name, displacement, fileLine);
        }

        static_cast<void>(SymCleanup(process));
        return out;
    }
    catch (const std::exception&)
    {
        // Best-effort: keep the crash pipeline moving; avoid allocations in the failure path.
        return {};
    }
}

[[nodiscard]] HRESULT WriteCrashReportFile(const std::filesystem::path& reportPath, EXCEPTION_POINTERS* exceptionPointers) noexcept
{
    if (reportPath.empty())
    {
        return E_INVALIDARG;
    }

    const std::wstring text = BuildStackTraceText(exceptionPointers);

    wil::unique_handle file(CreateFileW(reportPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const wchar_t bom = 0xFEFF;
    DWORD written     = 0;
    if (! WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const DWORD bytes = static_cast<DWORD>(text.size() * sizeof(wchar_t));
    if (bytes > 0)
    {
        if (! WriteFile(file.get(), text.data(), bytes, &written, nullptr))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }

    static_cast<void>(FlushFileBuffers(file.get()));
    return S_OK;
}

void WriteDumpAndMarker(EXCEPTION_POINTERS* exceptionPointers) noexcept
{
    // Mandatory: crash path `noexcept` boundary. Best-effort; do not throw from crash paths.
    try
    {
        const std::filesystem::path dir = GetCrashDirectory();
        if (dir.empty())
        {
            return;
        }

        const HRESULT hrDir = EnsureDirectoryExists(dir);
        if (FAILED(hrDir))
        {
            return;
        }

        const std::filesystem::path dumpPath = BuildDumpPath(dir);
        std::filesystem::path reportPath     = dumpPath;
        reportPath.replace_extension(L".txt");
        static_cast<void>(WriteCrashReportFile(reportPath, exceptionPointers));

        const HRESULT dumpHr = WriteMiniDumpFile(dumpPath, exceptionPointers);
        if (FAILED(dumpHr))
        {
            // Best-effort: even if the minidump fails, leave the report file behind for diagnostics.
            return;
        }

        const std::filesystem::path markerPath = GetCrashMarkerPath();
        static_cast<void>(WriteMarkerFile(markerPath, dumpPath.native()));
    }
    catch (const std::exception&)
    {
        // Best-effort: swallow standard exceptions and continue process termination.
    }
}

__declspec(nothrow) LONG WINAPI UnhandledExceptionFilterThunk(EXCEPTION_POINTERS* exceptionPointers)
{
    WriteDumpAndMarker(exceptionPointers);
    return EXCEPTION_EXECUTE_HANDLER;
}

void TerminateHandler() noexcept
{
    WriteDumpAndMarker(nullptr);
    TerminateProcess(GetCurrentProcess(), 1);
}

__declspec(nothrow) void __cdecl PureCallHandler()
{
    WriteDumpAndMarker(nullptr);
    TerminateProcess(GetCurrentProcess(), 1);
}

__declspec(nothrow) void __cdecl InvalidParameterHandler(
    const wchar_t* /*expression*/, const wchar_t* /*function*/, const wchar_t* /*file*/, unsigned int /*line*/, uintptr_t /*reserved*/)
{
    WriteDumpAndMarker(nullptr);
    TerminateProcess(GetCurrentProcess(), 1);
}

[[nodiscard]] std::wstring ReadMarkerDumpPath(const std::filesystem::path& markerPath) noexcept
{
    wil::unique_handle file(
        CreateFileW(markerPath.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return {};
    }

    LARGE_INTEGER size{};
    if (! GetFileSizeEx(file.get(), &size) || size.QuadPart <= 0)
    {
        return {};
    }

    const size_t bytes = static_cast<size_t>(std::min<LONGLONG>(size.QuadPart, 64 * 1024)); // cap
    std::wstring buffer;
    buffer.resize((bytes + sizeof(wchar_t) - 1) / sizeof(wchar_t));

    DWORD read = 0;
    if (! ReadFile(file.get(), buffer.data(), static_cast<DWORD>(buffer.size() * sizeof(wchar_t)), &read, nullptr) || read == 0)
    {
        return {};
    }

    buffer.resize(read / sizeof(wchar_t));
    if (! buffer.empty() && buffer.front() == 0xFEFF)
    {
        buffer.erase(buffer.begin());
    }

    // Trim trailing NUL/newlines/spaces.
    while (! buffer.empty())
    {
        const wchar_t ch = buffer.back();
        if (ch == L'\0' || ch == L'\r' || ch == L'\n' || ch == L' ' || ch == L'\t')
        {
            buffer.pop_back();
            continue;
        }
        break;
    }

    return buffer;
}
} // namespace

void Install() noexcept
{
    const bool wasInstalled = g_installed.exchange(true, std::memory_order_acq_rel);
    if (wasInstalled)
    {
        return;
    }

    SetUnhandledExceptionFilter(&UnhandledExceptionFilterThunk);
    std::set_terminate(&TerminateHandler);
    _set_purecall_handler(&PureCallHandler);
    _set_invalid_parameter_handler(&InvalidParameterHandler);
}

int WriteDumpForException(EXCEPTION_POINTERS* exceptionPointers) noexcept
{
    WriteDumpAndMarker(exceptionPointers);
    return EXCEPTION_EXECUTE_HANDLER;
}

void ShowPreviousCrashUiIfPresent(HWND ownerWindow) noexcept
{
    const std::filesystem::path markerPath = GetCrashMarkerPath();
    if (markerPath.empty())
    {
        return;
    }

    std::error_code ec;
    if (! std::filesystem::exists(markerPath, ec))
    {
        return;
    }

    const std::wstring dumpPath          = ReadMarkerDumpPath(markerPath);
    const std::filesystem::path crashDir = markerPath.parent_path();

    // Remove marker first to avoid repeated prompts if the user force-kills during the prompt.
    std::filesystem::remove(markerPath, ec);

    const std::wstring title = LoadStringResource(nullptr, IDS_CRASH_DETECTED_TITLE);
    const std::wstring message =
        dumpPath.empty() ? LoadStringResource(nullptr, IDS_CRASH_DETECTED_MESSAGE) : FormatStringResource(nullptr, IDS_CRASH_DETECTED_MESSAGE_FMT, dumpPath);

    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = (ownerWindow && IsWindow(ownerWindow)) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    prompt.severity      = HOST_ALERT_ERROR;
    prompt.buttons       = HOST_PROMPT_BUTTONS_YES_NO;
    prompt.targetWindow  = (prompt.scope == HOST_ALERT_SCOPE_WINDOW) ? ownerWindow : nullptr;
    prompt.title         = title.empty() ? nullptr : title.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_NO;

    HostPromptResult result = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt  = HostShowPrompt(prompt, nullptr, &result);
    if (FAILED(hrPrompt) || result != HOST_PROMPT_RESULT_YES)
    {
        return;
    }

    const std::wstring folder = crashDir.wstring();
    static_cast<void>(ShellExecuteW(ownerWindow, L"open", folder.c_str(), nullptr, nullptr, SW_SHOWNORMAL));
}

void TriggerCrashTest() noexcept
{
    // Non-continuable exception to validate the crash pipeline.
    RaiseException(0xE000CAFEu, EXCEPTION_NONCONTINUABLE, 0, nullptr);
}
} // namespace CrashHandler
