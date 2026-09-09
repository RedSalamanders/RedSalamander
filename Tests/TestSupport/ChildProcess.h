#pragma once

#include <Windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include <algorithm>
#include <array>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <format>
#include <new>
#include <stop_token>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <utility>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027)
namespace RedSalamander::TestSupport
{
enum class ChildProcessStatus : uint8_t
{
    LaunchFailed,
    Completed,
    TimedOut,
    Cancelled,
    WaitFailed,
    CaptureFailed,
};

struct ChildProcessOptions final
{
    std::filesystem::path executablePath;
    std::vector<std::wstring> arguments;
    std::filesystem::path workingDirectory;
    std::chrono::milliseconds timeout{30'000};
    size_t maxStdoutBytes = 16u * 1024u * 1024u;
    size_t maxStderrBytes = 16u * 1024u * 1024u;
    std::stop_token cancellationToken;
};

struct ChildProcessResult final
{
    ChildProcessStatus status = ChildProcessStatus::LaunchFailed;
    DWORD processId           = 0u;
    DWORD exitCode            = STILL_ACTIVE;
    DWORD win32Error          = ERROR_SUCCESS;
    std::string stdoutBytes;
    std::string stderrBytes;
    bool stdoutTruncated = false;
    bool stderrTruncated = false;
    std::chrono::milliseconds elapsed{};
    std::wstring diagnostic;

    [[nodiscard]] bool Completed() const noexcept
    {
        return status == ChildProcessStatus::Completed;
    }
};

[[nodiscard]] inline std::wstring QuoteWindowsCommandLineArgument(std::wstring_view argument)
{
    const bool needsQuotes = argument.empty() || argument.find_first_of(L" \t\n\v\"") != std::wstring_view::npos;
    if (! needsQuotes)
    {
        return std::wstring(argument);
    }

    std::wstring quoted;
    quoted.reserve(argument.size() + 2u);
    quoted.push_back(L'"');
    size_t pendingBackslashes = 0u;
    for (const wchar_t ch : argument)
    {
        if (ch == L'\\')
        {
            ++pendingBackslashes;
            continue;
        }

        if (ch == L'"')
        {
            quoted.append(pendingBackslashes * 2u + 1u, L'\\');
            quoted.push_back(L'"');
        }
        else
        {
            quoted.append(pendingBackslashes, L'\\');
            quoted.push_back(ch);
        }
        pendingBackslashes = 0u;
    }

    quoted.append(pendingBackslashes * 2u, L'\\');
    quoted.push_back(L'"');
    return quoted;
}

[[nodiscard]] inline std::wstring BuildWindowsCommandLine(const std::filesystem::path& executablePath, const std::vector<std::wstring>& arguments)
{
    std::wstring commandLine = QuoteWindowsCommandLineArgument(executablePath.wstring());
    for (const std::wstring& argument : arguments)
    {
        commandLine.push_back(L' ');
        commandLine.append(QuoteWindowsCommandLineArgument(argument));
    }
    return commandLine;
}

namespace ChildProcessDetail
{
struct PipeDrainState final
{
    std::string bytes;
    DWORD error    = ERROR_SUCCESS;
    bool truncated = false;
};

inline void AppendDiagnostic(std::wstring& diagnostic, std::wstring_view text)
{
    if (! diagnostic.empty())
    {
        diagnostic.append(L"; ");
    }
    diagnostic.append(text);
}

inline void SetFailure(ChildProcessResult& result, ChildProcessStatus status, std::wstring_view operation, DWORD error)
{
    result.status     = status;
    result.win32Error = error;
    result.diagnostic = std::format(L"{} failed. error={}", operation, error);
}

[[nodiscard]] inline bool CreateCapturedPipe(wil::unique_handle& readPipe, wil::unique_handle& writePipe, DWORD& outError) noexcept
{
    SECURITY_ATTRIBUTES securityAttributes{};
    securityAttributes.nLength        = sizeof(securityAttributes);
    securityAttributes.bInheritHandle = TRUE;
    if (CreatePipe(readPipe.put(), writePipe.put(), &securityAttributes, 0u) == FALSE)
    {
        outError = GetLastError();
        return false;
    }
    if (SetHandleInformation(readPipe.get(), HANDLE_FLAG_INHERIT, 0u) == FALSE)
    {
        outError = GetLastError();
        return false;
    }
    return true;
}

inline void DrainPipe(HANDLE pipe, size_t maximumBytes, PipeDrainState& state) noexcept
{
    try
    {
        std::array<char, 64u * 1024u> buffer{};
        for (;;)
        {
            DWORD bytesRead = 0u;
            if (ReadFile(pipe, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead, nullptr) == FALSE)
            {
                const DWORD error = GetLastError();
                if (error != ERROR_BROKEN_PIPE)
                {
                    state.error = error;
                }
                return;
            }
            if (bytesRead == 0u)
            {
                return;
            }

            const size_t available = maximumBytes > state.bytes.size() ? maximumBytes - state.bytes.size() : 0u;
            const size_t retained  = (std::min)(available, static_cast<size_t>(bytesRead));
            if (retained != 0u)
            {
                state.bytes.append(buffer.data(), retained);
            }
            if (retained != bytesRead)
            {
                state.truncated = true;
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory thread boundary: report capture failure without throwing out of the drain thread.
        state.error = ERROR_UNHANDLED_EXCEPTION;
    }
}

[[nodiscard]] inline ChildProcessResult RunChildProcessImpl(const ChildProcessOptions& options)
{
    ChildProcessResult result{};
    const auto startedAt = std::chrono::steady_clock::now();
    if (options.executablePath.empty())
    {
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"child executable validation", ERROR_INVALID_PARAMETER);
        return result;
    }

    wil::unique_handle stdoutRead;
    wil::unique_handle stdoutWrite;
    wil::unique_handle stderrRead;
    wil::unique_handle stderrWrite;
    DWORD pipeError = ERROR_SUCCESS;
    if (! CreateCapturedPipe(stdoutRead, stdoutWrite, pipeError) || ! CreateCapturedPipe(stderrRead, stderrWrite, pipeError))
    {
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"CreatePipe/SetHandleInformation", pipeError);
        return result;
    }

    SECURITY_ATTRIBUTES securityAttributes{};
    securityAttributes.nLength        = sizeof(securityAttributes);
    securityAttributes.bInheritHandle = TRUE;
    wil::unique_hfile nullInput(
        CreateFileW(L"NUL", GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, &securityAttributes, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! nullInput)
    {
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"CreateFileW(NUL)", GetLastError());
        return result;
    }

    wil::unique_handle job(CreateJobObjectW(nullptr, nullptr));
    if (! job)
    {
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"CreateJobObjectW", GetLastError());
        return result;
    }
    JOBOBJECT_EXTENDED_LIMIT_INFORMATION jobLimits{};
    jobLimits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
    if (SetInformationJobObject(job.get(), JobObjectExtendedLimitInformation, &jobLimits, sizeof(jobLimits)) == FALSE)
    {
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"SetInformationJobObject", GetLastError());
        return result;
    }

    SIZE_T attributeBytes = 0u;
    static_cast<void>(InitializeProcThreadAttributeList(nullptr, 1u, 0u, &attributeBytes));
    if (attributeBytes == 0u)
    {
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"InitializeProcThreadAttributeList(size)", GetLastError());
        return result;
    }
    std::vector<std::byte> attributeStorage(attributeBytes);
    auto* attributeList = reinterpret_cast<PPROC_THREAD_ATTRIBUTE_LIST>(attributeStorage.data());
    if (InitializeProcThreadAttributeList(attributeList, 1u, 0u, &attributeBytes) == FALSE)
    {
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"InitializeProcThreadAttributeList", GetLastError());
        return result;
    }
    const auto deleteAttributeList = wil::scope_exit([&] noexcept { DeleteProcThreadAttributeList(attributeList); });

    std::array<HANDLE, 3> inheritedHandles{nullInput.get(), stdoutWrite.get(), stderrWrite.get()};
    if (UpdateProcThreadAttribute(attributeList, 0u, PROC_THREAD_ATTRIBUTE_HANDLE_LIST, inheritedHandles.data(), sizeof(inheritedHandles), nullptr, nullptr) ==
        FALSE)
    {
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"UpdateProcThreadAttribute(HANDLE_LIST)", GetLastError());
        return result;
    }

    STARTUPINFOEXW startupInfo{};
    startupInfo.StartupInfo.cb         = sizeof(startupInfo);
    startupInfo.StartupInfo.dwFlags    = STARTF_USESTDHANDLES;
    startupInfo.StartupInfo.hStdInput  = nullInput.get();
    startupInfo.StartupInfo.hStdOutput = stdoutWrite.get();
    startupInfo.StartupInfo.hStdError  = stderrWrite.get();
    startupInfo.lpAttributeList        = attributeList;

    std::wstring executable  = options.executablePath.wstring();
    std::wstring commandLine = BuildWindowsCommandLine(options.executablePath, options.arguments);
    std::wstring workingDirectory;
    const wchar_t* workingDirectoryPointer = nullptr;
    if (! options.workingDirectory.empty())
    {
        workingDirectory        = options.workingDirectory.wstring();
        workingDirectoryPointer = workingDirectory.c_str();
    }

    PROCESS_INFORMATION processInfo{};
    if (CreateProcessW(executable.c_str(),
                       commandLine.data(),
                       nullptr,
                       nullptr,
                       TRUE,
                       CREATE_SUSPENDED | EXTENDED_STARTUPINFO_PRESENT,
                       nullptr,
                       workingDirectoryPointer,
                       &startupInfo.StartupInfo,
                       &processInfo) == FALSE)
    {
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"CreateProcessW", GetLastError());
        return result;
    }

    result.processId = processInfo.dwProcessId;
    wil::unique_handle process(processInfo.hProcess);
    wil::unique_handle primaryThread(processInfo.hThread);
    stdoutWrite.reset();
    stderrWrite.reset();
    nullInput.reset();

    if (AssignProcessToJobObject(job.get(), process.get()) == FALSE)
    {
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"AssignProcessToJobObject", GetLastError());
        static_cast<void>(TerminateProcess(process.get(), 1u));
        static_cast<void>(WaitForSingleObject(process.get(), 5000u));
        return result;
    }

    PipeDrainState stdoutState{};
    PipeDrainState stderrState{};
    std::jthread stdoutDrain;
    std::jthread stderrDrain;
    try
    {
        stdoutDrain = std::jthread([pipe = stdoutRead.get(), maximumBytes = options.maxStdoutBytes, &stdoutState]() noexcept
        { DrainPipe(pipe, maximumBytes, stdoutState); });
        stderrDrain = std::jthread([pipe = stderrRead.get(), maximumBytes = options.maxStderrBytes, &stderrState]() noexcept
        { DrainPipe(pipe, maximumBytes, stderrState); });
    }
    catch (const std::system_error&)
    {
        // Thread creation is required for deadlock-free concurrent drains; fail launch and contain the suspended child.
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"stdout/stderr drain thread creation", ERROR_NOT_ENOUGH_MEMORY);
        job.reset();
        static_cast<void>(TerminateProcess(process.get(), 1u));
        static_cast<void>(WaitForSingleObject(process.get(), 5000u));
        if (stdoutDrain.joinable())
        {
            stdoutDrain.join();
        }
        return result;
    }

    if (ResumeThread(primaryThread.get()) == static_cast<DWORD>(-1))
    {
        SetFailure(result, ChildProcessStatus::LaunchFailed, L"ResumeThread", GetLastError());
        job.reset();
        static_cast<void>(TerminateProcess(process.get(), 1u));
        static_cast<void>(WaitForSingleObject(process.get(), 5000u));
    }
    else
    {
        result.status       = ChildProcessStatus::Completed;
        const auto deadline = startedAt + options.timeout;
        for (;;)
        {
            if (options.cancellationToken.stop_requested())
            {
                result.status     = ChildProcessStatus::Cancelled;
                result.diagnostic = L"Child process execution was cancelled.";
                break;
            }

            const auto now = std::chrono::steady_clock::now();
            if (now >= deadline)
            {
                const DWORD finalProbe = WaitForSingleObject(process.get(), 0u);
                if (finalProbe == WAIT_OBJECT_0)
                {
                    break;
                }
                if (finalProbe == WAIT_FAILED)
                {
                    SetFailure(result, ChildProcessStatus::WaitFailed, L"WaitForSingleObject", GetLastError());
                }
                else
                {
                    result.status     = ChildProcessStatus::TimedOut;
                    result.diagnostic = std::format(L"Child process timed out after {} ms.", options.timeout.count());
                }
                break;
            }

            const auto remaining   = std::chrono::duration_cast<std::chrono::milliseconds>(deadline - now);
            const DWORD waitSlice  = static_cast<DWORD>((std::max)(1ll, (std::min)(50ll, remaining.count())));
            const DWORD waitResult = WaitForSingleObject(process.get(), waitSlice);
            if (waitResult == WAIT_OBJECT_0)
            {
                break;
            }
            if (waitResult == WAIT_FAILED)
            {
                SetFailure(result, ChildProcessStatus::WaitFailed, L"WaitForSingleObject", GetLastError());
                break;
            }
        }
    }

    // Closing a kill-on-close job after the root exits also contains any descendant that tried to outlive it.
    job.reset();
    if (result.status != ChildProcessStatus::Completed)
    {
        static_cast<void>(TerminateProcess(process.get(), 1u));
        static_cast<void>(WaitForSingleObject(process.get(), 5000u));
    }
    if (stdoutDrain.joinable())
    {
        stdoutDrain.join();
    }
    if (stderrDrain.joinable())
    {
        stderrDrain.join();
    }

    result.stdoutBytes     = std::move(stdoutState.bytes);
    result.stderrBytes     = std::move(stderrState.bytes);
    result.stdoutTruncated = stdoutState.truncated;
    result.stderrTruncated = stderrState.truncated;
    if (stdoutState.error != ERROR_SUCCESS || stderrState.error != ERROR_SUCCESS)
    {
        const DWORD captureError = stdoutState.error != ERROR_SUCCESS ? stdoutState.error : stderrState.error;
        if (result.status == ChildProcessStatus::Completed)
        {
            SetFailure(result, ChildProcessStatus::CaptureFailed, L"stdout/stderr pipe drain", captureError);
        }
        else
        {
            AppendDiagnostic(result.diagnostic, std::format(L"pipe drain error={}", captureError));
        }
    }

    if (result.status == ChildProcessStatus::Completed && GetExitCodeProcess(process.get(), &result.exitCode) == FALSE)
    {
        SetFailure(result, ChildProcessStatus::WaitFailed, L"GetExitCodeProcess", GetLastError());
    }
    result.elapsed = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - startedAt);
    return result;
}
} // namespace ChildProcessDetail

[[nodiscard]] inline ChildProcessResult RunChildProcess(const ChildProcessOptions& options) noexcept
{
    try
    {
        return ChildProcessDetail::RunChildProcessImpl(options);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::system_error&)
    {
        // Mandatory noexcept test-infrastructure boundary: return a launch diagnostic for runtime failures.
        ChildProcessResult result{};
        ChildProcessDetail::SetFailure(result, ChildProcessStatus::LaunchFailed, L"child process runtime", ERROR_UNHANDLED_EXCEPTION);
        return result;
    }
    catch (const std::exception&)
    {
        // Mandatory noexcept test-infrastructure boundary: preserve a deterministic failure instead of unwinding a selftest.
        ChildProcessResult result{};
        ChildProcessDetail::SetFailure(result, ChildProcessStatus::LaunchFailed, L"child process setup", ERROR_UNHANDLED_EXCEPTION);
        return result;
    }
}
} // namespace RedSalamander::TestSupport
#pragma warning(pop)
