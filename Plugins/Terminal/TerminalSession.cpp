#include "Terminal.h"
#include "TerminalAccessibility.h"
#include "TerminalInternal.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <clocale>
#include <cmath>
#include <cstdlib>
#include <cstring>
#include <format>
#include <limits>
#include <new>
#include <numeric>
#include <optional>
#include <ranges>
#include <set>
#include <system_error>
#include <utility>
#include <vector>

#include <windowsx.h>
#include <wincrypt.h>
#include <imm.h>
#include <shlobj.h>
#include <shellapi.h>
#include <UIAutomation.h>
#include <uxtheme.h>

#pragma comment(lib, "crypt32")
#pragma comment(lib, "imm32")
#pragma comment(lib, "advapi32")
#pragma comment(lib, "shell32")
#pragma comment(lib, "uxtheme")

#include <yyjson.h>

#include "Helpers.h"
#include "DxUi/DxUi.h"
#include "PathUtils.h"
#include "PaneVisualState.h"
#include "ProcessCommandLine.h"
#include "StringConversion.h"
#include "UnicodeClipboard.h"
#include "WindowMessages.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

using namespace TerminalPluginDetail;

#include "TerminalSession.h"

HRESULT Terminal::resolveShell(const TerminalLogicalLocation& location,
                               std::wstring& executable,
                               std::wstring& commandLine,
                               ShellKind& kind) noexcept
{
    if (location.kind == TerminalLocationKind::Wsl)
    {
        executable = GetSystemExecutable(L"wsl.exe");
        if (executable.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
        const std::wstring distribution = copySpan(location.wslDistribution);
        const std::wstring path = copySpan(location.wslAbsolutePath);
        if (distribution.empty() || distribution.front() == L'-' || distribution.find_first_of(L"\\/") != std::wstring::npos || path.empty() ||
            path.front() != L'/' || ContainsCommandControl(distribution) || ContainsCommandControl(path))
        {
            return E_INVALIDARG;
        }
        kind = ShellKind::Posix;
        commandLine = std::format(L"{} --distribution {} --cd {}",
                                  Common::Process::QuoteWindowsCommandLineArgument(executable),
                                  Common::Process::QuoteWindowsCommandLineArgument(distribution),
                                  Common::Process::QuoteWindowsCommandLineArgument(path));
        if (commandLine.size() >= 32767u)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        return S_OK;
    }

    if (_config.defaultShell != L"cmd")
    {
        if (_config.defaultShell == L"auto")
        {
            executable = FindPowerShell7();
        }
        if (executable.empty())
        {
            executable = GetSystemExecutable(L"WindowsPowerShell\\v1.0\\powershell.exe");
        }
        if (! executable.empty())
        {
            kind = ShellKind::PowerShell;
            commandLine = Common::Process::QuoteWindowsCommandLineArgument(executable) + L" -NoLogo";
            return S_OK;
        }
        if (_config.defaultShell == L"powershell")
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
    }

    executable = GetSystemExecutable(L"cmd.exe");
    if (executable.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    kind = ShellKind::CommandPrompt;
    commandLine = Common::Process::QuoteWindowsCommandLineArgument(executable) + L" /D /Q /V:OFF";
    return S_OK;
}

HRESULT Terminal::startPseudoConsole(const TerminalLogicalLocation& location) noexcept
{
    wil::unique_handle inputRead;
    wil::unique_handle outputRead;
    wil::unique_handle outputWrite;
    if (CreatePipe(inputRead.put(), _ptyInputWrite.put(), nullptr, 0u) == FALSE ||
        CreatePipe(outputRead.put(), outputWrite.put(), nullptr, 0u) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    _ptyClientSizeSynced = false;
    HPCON rawPseudoConsole = nullptr;
    HRESULT hr = CreatePseudoConsole(COORD{static_cast<SHORT>(_columns), static_cast<SHORT>(_rows)},
                                     inputRead.get(),
                                     outputWrite.get(),
                                     0u,
                                     &rawPseudoConsole);
    if (FAILED(hr))
    {
        return hr;
    }
    _pseudoConsole.reset(rawPseudoConsole);
    inputRead.reset();
    outputWrite.reset();

    SIZE_T attributeBytes = 0u;
    static_cast<void>(InitializeProcThreadAttributeList(nullptr, 1u, 0u, &attributeBytes));
    if (attributeBytes == 0u)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    std::vector<std::byte> attributeStorage(attributeBytes);
    auto* attributes = reinterpret_cast<PPROC_THREAD_ATTRIBUTE_LIST>(attributeStorage.data());
    if (InitializeProcThreadAttributeList(attributes, 1u, 0u, &attributeBytes) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    const auto deleteAttributes = wil::scope_exit([&] noexcept { DeleteProcThreadAttributeList(attributes); });
    const HPCON pseudoConsole = _pseudoConsole.get();
    // This attribute takes the HPCON value itself. Passing &pseudoConsole gives
    // the hosted process an invalid pseudoconsole and produces 0xc0000142.
    if (UpdateProcThreadAttribute(attributes,
                                  0u,
                                  PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE,
                                  pseudoConsole,
                                  sizeof(pseudoConsole),
                                  nullptr,
                                  nullptr) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    std::wstring executable;
    std::wstring commandLine;
    HRESULT shellHr = resolveShell(location, executable, commandLine, _shellKind);
    if (FAILED(shellHr))
    {
        return shellHr;
    }

    std::wstring integrationPipeLeaf;
    std::wstring integrationPipeName;
    if (_shellKind == ShellKind::PowerShell)
    {
        RETURN_IF_FAILED(MakeIntegrationPipeName(integrationPipeLeaf, integrationPipeName));
        std::wstring encodedIntegration;
        RETURN_IF_FAILED(EncodePowerShellCommand(integrationPipeLeaf, encodedIntegration));
        commandLine.append(L" -NoExit");
        // Keep an explicit interactive session after the bootstrap command.
        // Windows PowerShell 5.1 does not accept this switch.
        if (executable.size() >= 8u &&
            OrdinalString::EqualsNoCase(std::wstring_view(executable).substr(executable.size() - 8u), L"pwsh.exe"))
        {
            commandLine.append(L" -Interactive");
        }
        commandLine.append(L" -EncodedCommand ");
        commandLine.append(encodedIntegration);
        if (commandLine.size() >= 32767u)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
    }

    wil::unique_handle processJob(CreateJobObjectW(nullptr, nullptr));
    if (! processJob)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    JOBOBJECT_EXTENDED_LIMIT_INFORMATION jobLimits{};
    jobLimits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
    if (SetInformationJobObject(
            processJob.get(), JobObjectExtendedLimitInformation, &jobLimits, sizeof(jobLimits)) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    STARTUPINFOEXW startup{};
    startup.StartupInfo.cb = sizeof(startup);
    // Explicit null standard handles let ConPTY supply all three streams.
    // Otherwise Windows may duplicate redirected host handles even with
    // bInheritHandles=FALSE. CREATE_NO_WINDOW instead detaches console I/O.
    startup.StartupInfo.dwFlags = STARTF_USESTDHANDLES;
    startup.lpAttributeList = attributes;
    PROCESS_INFORMATION processInfo{};
    std::wstring ownedWorkingDirectory;
    if (location.kind == TerminalLocationKind::WindowsLocal || location.kind == TerminalLocationKind::WindowsUnc)
    {
        ownedWorkingDirectory = copySpan(location.windowsPath);
    }
    if (CreateProcessW(executable.c_str(),
                       commandLine.data(),
                       nullptr,
                       nullptr,
                       FALSE,
                       EXTENDED_STARTUPINFO_PRESENT | CREATE_UNICODE_ENVIRONMENT | CREATE_SUSPENDED,
                       nullptr,
                       ownedWorkingDirectory.empty() ? nullptr : ownedWorkingDirectory.c_str(),
                       &startup.StartupInfo,
                       &processInfo) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    wil::unique_handle process(processInfo.hProcess);
    wil::unique_handle processThread(processInfo.hThread);
    auto terminateSuspendedProcess = wil::scope_exit([&] noexcept
    {
        if (process)
        {
            static_cast<void>(TerminateProcess(process.get(), ERROR_CANCELLED));
        }
    });
    if (AssignProcessToJobObject(processJob.get(), process.get()) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    wil::unique_handle rootProcessForWatcher;
    if (DuplicateHandle(GetCurrentProcess(),
                        process.get(),
                        GetCurrentProcess(),
                        rootProcessForWatcher.put(),
                        0u,
                        FALSE,
                        DUPLICATE_SAME_ACCESS) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    std::jthread integrationThread;
    if (! integrationPipeName.empty())
    {
        try
        {
            integrationThread = std::jthread(
                [this, rootProcessId = processInfo.dwProcessId, pipeName = std::move(integrationPipeName)](
                    std::stop_token stopToken) mutable noexcept
                {
                    integrationMain(stopToken, rootProcessId, std::move(pipeName));
                });
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
        catch (const std::system_error& error)
        {
            return HRESULT_FROM_WIN32(static_cast<DWORD>(error.code().value()));
        }
    }
    if (ResumeThread(processThread.get()) == static_cast<DWORD>(-1))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    // Keep the rollback owner armed after resume. A failed reader-thread start
    // must still terminate the now-running hidden shell before Open returns.

    try
    {
        std::jthread readerThread([this, outputRead = std::move(outputRead)](std::stop_token stopToken) mutable noexcept
        {
            readerMain(stopToken, std::move(outputRead));
        });
        _integrationThread = std::move(integrationThread);
        _readerThread = std::move(readerThread);
        _processJob = std::move(processJob);
        _process = std::move(process);
        _processWatcherThread = std::jthread(
            [this, rootProcess = std::move(rootProcessForWatcher)](std::stop_token stopToken) mutable noexcept
            {
                processWatcherMain(stopToken, std::move(rootProcess));
            });
        terminateSuspendedProcess.release();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::system_error& error)
    {
        _processJob.reset();
        if (_process)
        {
            static_cast<void>(TerminateProcess(_process.get(), ERROR_CANCELLED));
        }
        {
            std::scoped_lock lock(_ioMutex);
            _ptyInputWrite.reset();
        }
        _pseudoConsole.reset();
        if (_readerThread.joinable())
        {
            _readerThread.request_stop();
            static_cast<void>(CancelSynchronousIo(_readerThread.native_handle()));
            _readerThread.join();
        }
        if (_integrationThread.joinable())
        {
            _integrationThread.request_stop();
            _integrationThread.join();
        }
        _process.reset();
        return HRESULT_FROM_WIN32(static_cast<DWORD>(error.code().value()));
    }
    return S_OK;
}

void Terminal::integrationMain(std::stop_token stopToken, DWORD rootProcessId, std::wstring pipeName) noexcept
{
#if defined(ENABLE_TESTS)
    _integrationDebugStage.store(1u, std::memory_order_release);
#endif
    while (! stopToken.stop_requested() && ! _closing.load(std::memory_order_acquire))
    {
        wil::unique_handle pipe(CreateNamedPipeW(pipeName.c_str(),
                                                 PIPE_ACCESS_DUPLEX | FILE_FLAG_OVERLAPPED | FILE_FLAG_FIRST_PIPE_INSTANCE,
                                                 PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT | PIPE_REJECT_REMOTE_CLIENTS,
                                                 1u,
                                                 static_cast<DWORD>(kMaximumIntegrationPathBytes),
                                                 static_cast<DWORD>(kMaximumIntegrationPathBytes),
                                                 0u,
                                                 nullptr));
        if (! pipe)
        {
#if defined(ENABLE_TESTS)
            _integrationDebugResult.store(HRESULT_FROM_WIN32(GetLastError()), std::memory_order_release);
#endif
            return;
        }
#if defined(ENABLE_TESTS)
        _integrationDebugStage.store(2u, std::memory_order_release);
#endif
        const HRESULT connectHr = ConnectIntegrationPipe(pipe.get(), stopToken);
        if (FAILED(connectHr))
        {
#if defined(ENABLE_TESTS)
            _integrationDebugResult.store(connectHr, std::memory_order_release);
#endif
            if (connectHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return;
            }
            continue;
        }
#if defined(ENABLE_TESTS)
        _integrationDebugStage.store(3u, std::memory_order_release);
#endif

        ULONG clientProcessId = 0u;
#if defined(ENABLE_TESTS)
        _integrationDebugClientPid.store(clientProcessId, std::memory_order_release);
#endif
        if (GetNamedPipeClientProcessId(pipe.get(), &clientProcessId) == FALSE || clientProcessId != rootProcessId)
        {
#if defined(ENABLE_TESTS)
            _integrationDebugClientPid.store(clientProcessId, std::memory_order_release);
            _integrationDebugResult.store(
                clientProcessId != rootProcessId ? E_ACCESSDENIED : HRESULT_FROM_WIN32(GetLastError()),
                std::memory_order_release);
#endif
            static_cast<void>(DisconnectNamedPipe(pipe.get()));
            continue;
        }
#if defined(ENABLE_TESTS)
        _integrationDebugStage.store(4u, std::memory_order_release);
        _integrationDebugClientPid.store(clientProcessId, std::memory_order_release);
#endif
        const auto disconnectPipe = wil::scope_exit([&] noexcept { static_cast<void>(DisconnectNamedPipe(pipe.get())); });
        std::string utf8;
        const HRESULT readHr = ReadIntegrationPath(pipe.get(), stopToken, utf8);
        if (FAILED(readHr))
        {
#if defined(ENABLE_TESTS)
            _integrationDebugResult.store(readHr, std::memory_order_release);
#endif
            if (readHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return;
            }
            continue;
        }
#if defined(ENABLE_TESTS)
        _integrationDebugStage.store(5u, std::memory_order_release);
#endif
        IntegrationRequest request;
        const HRESULT parseHr = ParseIntegrationRequest(utf8, request);
        SecureZeroMemory(utf8.data(), utf8.size());
        utf8.clear();
        const auto scrubRequest = wil::scope_exit([&request]() noexcept
        {
            SecureZeroMemory(request.promptBuffer.data(), request.promptBuffer.size() * sizeof(wchar_t));
            ScrubStrings(request.currentSessionHistory);
        });
        if (FAILED(parseHr) || _closing.load(std::memory_order_acquire))
        {
#if defined(ENABLE_TESTS)
            _integrationDebugResult.store(FAILED(parseHr) ? parseHr : HRESULT_FROM_WIN32(ERROR_CANCELLED),
                                          std::memory_order_release);
#endif
            continue;
        }
#if defined(ENABLE_TESTS)
        _integrationDebugStage.store(6u, std::memory_order_release);
        _integrationDebugResult.store(S_OK, std::memory_order_release);
#endif
        TrustedPromptReply reply = acceptTrustedPrompt(std::move(request.currentDirectory),
                                                       std::move(request.historyPath),
                                                       std::move(request.promptBuffer),
                                                       std::move(request.currentSessionHistory),
                                                       request.versioned);
        const auto scrubReply = wil::scope_exit([&reply]() noexcept
        {
            SecureZeroMemory(reply.promptReplacement.data(), reply.promptReplacement.size() * sizeof(wchar_t));
        });
        const HRESULT writeHr = WriteIntegrationReply(
            pipe.get(), stopToken, request.versioned, reply.followTarget, reply.promptReplacement);
        if (FAILED(writeHr))
        {
#if defined(ENABLE_TESTS)
            _integrationDebugResult.store(writeHr, std::memory_order_release);
#endif
            if (! reply.followTarget.empty())
            {
                {
                    std::scoped_lock stateLock(_stateMutex);
                    _followState = TerminalFollowState::Failed;
                    _idleAtPrimaryPrompt = true;
                    _commandSubmitted = false;
                    _stateGeneration.fetch_add(1u, std::memory_order_acq_rel);
                }
            }
            if (reply.replacementAccepted)
            {
                reply.replacementAccepted = false;
                reply.replacementRejected = true;
            }
            if (writeHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return;
            }
        }
        if (reply.replacementAccepted || reply.replacementRejected)
        {
            if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
            {
                static_cast<void>(PostMessageW(hwnd,
                                               WndMsg::kTerminalSuggestionInsertionComplete,
                                               reply.replacementAccepted ? TRUE : FALSE,
                                               0));
            }
        }
        if (FAILED(writeHr))
        {
            continue;
        }
#if defined(ENABLE_TESTS)
        _integrationDebugStage.store(7u, std::memory_order_release);
#endif
    }
}

Terminal::TrustedPromptReply Terminal::acceptTrustedPrompt(std::wstring path,
                                                           std::wstring historyPath,
                                                           std::wstring promptBuffer,
                                                           std::vector<std::wstring> currentSessionHistory,
                                                           bool supportsPromptReplacement) noexcept
{
    TrustedPromptReply reply;
    const bool followEnabled = _followPathWhenIdle.load(std::memory_order_acquire);
    {
        std::scoped_lock stateLock(_stateMutex);
        const TerminalFollowState previousFollowState = _followState;
        const bool stateChanged = ! _integrationTrusted ||
            _trustedHistoryProviderAvailable != supportsPromptReplacement || ! _idleAtPrimaryPrompt || _commandSubmitted ||
            ! OrdinalString::EqualsNoCase(std::wstring_view(_trustedCurrentDirectory), std::wstring_view(path)) ||
            ! OrdinalString::EqualsNoCase(std::wstring_view(_trustedHistoryPath), std::wstring_view(historyPath)) ||
            _trustedPromptBuffer != promptBuffer || _trustedCurrentSessionHistory != currentSessionHistory;
        _integrationTrusted = true;
        _trustedHistoryProviderAvailable = supportsPromptReplacement;
        _idleAtPrimaryPrompt = true;
        _commandSubmitted = false;
        _hasPendingUserInput = ! promptBuffer.empty();
        _trustedCurrentDirectory = std::move(path);
        _trustedHistoryPath = std::move(historyPath);
        SecureZeroMemory(_trustedPromptBuffer.data(), _trustedPromptBuffer.size() * sizeof(wchar_t));
        _trustedPromptBuffer = std::move(promptBuffer);
        ScrubStrings(_trustedCurrentSessionHistory);
        _trustedCurrentSessionHistory = std::move(currentSessionHistory);

        const bool replacementPending = ! _pendingPromptReplacement.empty();
        if (replacementPending)
        {
            if (supportsPromptReplacement && _pendingPromptExpectedBuffer == _trustedPromptBuffer)
            {
                reply.promptReplacement = _pendingPromptReplacement;
                reply.replacementAccepted = true;
            }
            else
            {
                reply.replacementRejected = true;
            }
            SecureZeroMemory(
                _pendingPromptExpectedBuffer.data(), _pendingPromptExpectedBuffer.size() * sizeof(wchar_t));
            SecureZeroMemory(_pendingPromptReplacement.data(), _pendingPromptReplacement.size() * sizeof(wchar_t));
            _pendingPromptExpectedBuffer.clear();
            _pendingPromptReplacement.clear();
        }
        if (followEnabled && ! _pendingFollowDirectory.empty() &&
            OrdinalString::EqualsNoCase(
                std::wstring_view(_trustedCurrentDirectory), std::wstring_view(_pendingFollowDirectory)))
        {
            _followState = TerminalFollowState::Applied;
            _pendingFollowDirectory.clear();
        }
        else if (! replacementPending && followEnabled && ! _hasPendingUserInput && ! _pendingFollowDirectory.empty())
        {
            reply.followTarget = _pendingFollowDirectory;
            _followState = TerminalFollowState::Applying;
            _idleAtPrimaryPrompt = false;
            _commandSubmitted = true;
        }
        else if (followEnabled && _pendingFollowDirectory.empty())
        {
            _followState = TerminalFollowState::Applied;
        }
        if (stateChanged || previousFollowState != _followState || ! reply.followTarget.empty() || replacementPending)
        {
            _stateGeneration.fetch_add(1u, std::memory_order_acq_rel);
        }
    }
    return reply;
}

void Terminal::processWatcherMain(std::stop_token stopToken, wil::unique_handle rootProcess) noexcept
{
    while (! stopToken.stop_requested())
    {
        const DWORD waitResult = WaitForSingleObject(rootProcess.get(), 50u);
        if (waitResult == WAIT_TIMEOUT)
        {
            continue;
        }
        if (waitResult != WAIT_OBJECT_0)
        {
            if (waitResult == WAIT_FAILED)
            {
                Debug::ErrorWithLastError(L"Terminal: waiting for the root process failed.");
            }
            return;
        }
        break;
    }
    if (stopToken.stop_requested() || _closing.load(std::memory_order_acquire))
    {
        return;
    }

    DWORD exitCode = 0u;
    if (GetExitCodeProcess(rootProcess.get(), &exitCode) == FALSE || exitCode == STILL_ACTIVE)
    {
        return;
    }
    _rootExitObservedTimestampNs.store(SteadyTimestampNs(), std::memory_order_release);
    _exitCode.store(exitCode, std::memory_order_release);
    _exitCodePresent.store(true, std::memory_order_release);
    _rootProcessExited.store(true, std::memory_order_release);
    setLifecycle(TerminalLifecycleState::DrainingAfterRootExit);

    std::scoped_lock teardownLock(_sessionTeardownMutex);
    if (_closing.load(std::memory_order_acquire))
    {
        return;
    }
    // The root shell owns the session lifetime. Retire its remaining job tree,
    // stop accepting input, then close HPCON while readerMain continues draining
    // the final ConPTY output to EOF (required on pre-26100 systems).
    _processJob.reset();
    {
        std::scoped_lock ioLock(_ioMutex);
        _ptyInputWrite.reset();
    }
    _pseudoConsole.reset();
}

void Terminal::readerMain(std::stop_token stopToken, wil::unique_handle outputRead) noexcept
{
    const TerminalPngDecodeThreadContext pngDecodeContext;
    if (const HRESULT pngStatus = pngDecodeContext.Status(); FAILED(pngStatus))
    {
        Debug::Warning(L"Terminal PNG decoder is unavailable on the VT reader thread (hr=0x{:08X}).",
                       static_cast<uint32_t>(pngStatus));
    }
    std::array<uint8_t, 16u * 1024u> buffer{};
    // During normal close this reader deliberately continues to EOF. Microsoft
    // documents that pre-26100 ClosePseudoConsole can wait for the output pipe
    // to be drained; the close worker kills the child tree and closes HPCON
    // while this thread remains the single drain owner. stop_token is reserved
    // for the post-HPCON cancellation fallback.
    while (! stopToken.stop_requested())
    {
        DWORD bytesRead = 0u;
        const HANDLE output = outputRead.get();
        if (output == nullptr || ReadFile(output, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead, nullptr) == FALSE || bytesRead == 0u)
        {
            break;
        }
        {
            std::scoped_lock lock(_terminalMutex);
            if (_ghosttyTerminal != nullptr)
            {
                _runtime.terminalVtWrite(_ghosttyTerminal, buffer.data(), bytesRead);
            }
        }
        _lastOutputTimestampNs.store(SteadyTimestampNs(), std::memory_order_release);
        notifyOutputReady();
    }
    if (! _closing.load(std::memory_order_acquire) && _rootProcessExited.load(std::memory_order_acquire))
    {
        _finalSnapshotComplete.store(true, std::memory_order_release);
        setLifecycle(TerminalLifecycleState::Exited);
        notifyOutputReady();
    }
}

void Terminal::notifyOutputReady() noexcept
{
    _outputGeneration.fetch_add(1u, std::memory_order_acq_rel);
    if (_closing.load(std::memory_order_acquire))
    {
        return;
    }

    bool expected = false;
    if (! _outputNotificationPending.compare_exchange_strong(
            expected, true, std::memory_order_acq_rel, std::memory_order_acquire))
    {
        return;
    }
    postOutputReadyWake();
}

void Terminal::postOutputReadyWake() noexcept
{
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    auto payload = std::unique_ptr<ReadyWakePayload>(new (std::nothrow) ReadyWakePayload());
    if (hwnd == nullptr || _closing.load(std::memory_order_acquire) || ! payload)
    {
        bool expected = true;
        static_cast<void>(_outputNotificationPending.compare_exchange_strong(
            expected, false, std::memory_order_acq_rel, std::memory_order_acquire));
        return;
    }
    payload->sessionGeneration = _sessionGeneration.load(std::memory_order_acquire);
    if (! PostMessagePayload(hwnd, WndMsg::kTerminalOutputReady, 0u, std::move(payload)))
    {
        bool expected = true;
        static_cast<void>(_outputNotificationPending.compare_exchange_strong(
            expected, false, std::memory_order_acq_rel, std::memory_order_acquire));
    }
}

void Terminal::consumeOutputReady() noexcept
{
    const uint64_t consumed = _outputGeneration.load(std::memory_order_acquire);
    _outputConsumedGeneration.store(consumed, std::memory_order_release);
    _outputNotificationPending.store(false, std::memory_order_release);

    // A producer can advance the generation after the snapshot but before the
    // pending flag is cleared. Re-arm through the same false-to-true gate so the
    // newest generation cannot be stranded in that handoff.
    if (_outputGeneration.load(std::memory_order_acquire) != consumed)
    {
        bool expected = false;
        if (! _closing.load(std::memory_order_acquire) &&
            _outputNotificationPending.compare_exchange_strong(
                expected, true, std::memory_order_acq_rel, std::memory_order_acquire))
        {
            postOutputReadyWake();
        }
    }
    publishFinalExitEvent();
}

void Terminal::publishFinalExitEvent() noexcept
{
    if (_closing.load(std::memory_order_acquire) || ! _finalSnapshotComplete.load(std::memory_order_acquire) ||
        _lifecycle.load(std::memory_order_acquire) != TerminalLifecycleState::Exited)
    {
        return;
    }

    const uint64_t sessionGeneration = _sessionGeneration.load(std::memory_order_acquire);
    std::scoped_lock callbackLock(_eventCallbackMutex);
    if (_eventCallback == nullptr || _eventDeliveredSessionGeneration == sessionGeneration)
    {
        return;
    }

    TerminalEvent event{};
    event.sizeBytes = sizeof(event);
    event.kind = TerminalEventKind::RootSessionExited;
    event.instanceId = _instanceId;
    event.sessionGeneration = sessionGeneration;
    event.stateGeneration = _stateGeneration.load(std::memory_order_acquire);
    event.exitCodePresent = _exitCodePresent.load(std::memory_order_acquire) ? 1u : 0u;
    event.exitCode = _exitCode.load(std::memory_order_acquire);
    event.finalSnapshotComplete = 1u;
    event.rootExitObservedTimestampNs = _rootExitObservedTimestampNs.load(std::memory_order_acquire);
    _eventDeliveredSessionGeneration = sessionGeneration;
    _eventCallback->OnTerminalEvent(&event, _eventCallbackCookie);
}

void Terminal::GhosttyWritePty(GhosttyTerminal /*terminal*/, void* userData, const uint8_t* data, size_t length) noexcept
{
    if (userData != nullptr && data != nullptr && length != 0u)
    {
        static_cast<void>(
            static_cast<Terminal*>(userData)->writeInput(std::string_view(reinterpret_cast<const char*>(data), length)));
    }
}

bool Terminal::writeInput(std::string_view bytes) noexcept
{
    if (bytes.empty())
    {
        return true;
    }
    if (bytes.size() > kMaximumInputBytes)
    {
        return false;
    }

    std::unique_lock lock(_inputMutex);
    if (_closing.load(std::memory_order_acquire) ||
        _inputQueue.size() + _inputActiveRequests >= kMaximumInputRequests ||
        _inputQueuedBytes > kMaximumInputBytes - bytes.size())
    {
        return false;
    }
    _inputQueue.emplace_back(bytes);
    _inputQueuedBytes += bytes.size();
    if (_inputDrainScheduled)
    {
#if defined(ENABLE_TESTS)
        _debugCapturedInput.append(bytes);
#endif
        return true;
    }
    _inputDrainScheduled = true;

    wil::unique_hmodule modulePin = AcquireModuleReferenceFromAddress(&kTerminalModuleAnchor);
    auto context = modulePin ?
        std::unique_ptr<InputDrainContext>(new (std::nothrow) InputDrainContext(this, std::move(modulePin))) : nullptr;
    if (! context)
    {
        _inputDrainScheduled = false;
        _inputQueuedBytes -= _inputQueue.back().size();
        SecureZeroMemory(_inputQueue.back().data(), _inputQueue.back().size());
        _inputQueue.pop_back();
        return false;
    }

    AddRef();
    if (TrySubmitThreadpoolCallback(&Terminal::InputDrainCallback, context.get(), nullptr) == FALSE)
    {
        Release();
        _inputDrainScheduled = false;
        _inputQueuedBytes -= _inputQueue.back().size();
        SecureZeroMemory(_inputQueue.back().data(), _inputQueue.back().size());
        _inputQueue.pop_back();
        return false;
    }
    static_cast<void>(context.release());
#if defined(ENABLE_TESTS)
    _debugCapturedInput.append(bytes);
#endif
    return true;
}

void CALLBACK Terminal::InputDrainCallback(PTP_CALLBACK_INSTANCE instance, void* context) noexcept
{
    std::unique_ptr<InputDrainContext> owned(static_cast<InputDrainContext*>(context));
    if (! owned || owned->terminal == nullptr)
    {
        return;
    }
    Terminal* terminal = owned->terminal;
    TransferModulePinToCallbackReturn(instance, owned->modulePin);
    owned.reset();
    terminal->drainInput();
    terminal->Release();
}

void Terminal::drainInput() noexcept
{
    for (;;)
    {
        std::string bytes;
        {
            std::scoped_lock lock(_inputMutex);
            if (_closing.load(std::memory_order_acquire) || _inputQueue.empty())
            {
                _inputDrainScheduled = false;
                return;
            }
            bytes = std::move(_inputQueue.front());
            _inputQueue.pop_front();
            ++_inputActiveRequests;
        }

        {
            std::scoped_lock lock(_ioMutex);
            size_t offset = 0u;
            while (_ptyInputWrite && offset < bytes.size())
            {
                const size_t remaining = bytes.size() - offset;
                const DWORD requested = static_cast<DWORD>(
                    std::min<size_t>(remaining, (std::numeric_limits<DWORD>::max)()));
                DWORD written = 0u;
                if (WriteFile(_ptyInputWrite.get(), bytes.data() + offset, requested, &written, nullptr) == FALSE || written == 0u)
                {
                    break;
                }
                offset += written;
            }
        }

        {
            std::scoped_lock lock(_inputMutex);
            _inputQueuedBytes -= std::min(_inputQueuedBytes, bytes.size());
            --_inputActiveRequests;
        }
        SecureZeroMemory(bytes.data(), bytes.size());
    }
}

void Terminal::clearQueuedInput() noexcept
{
    std::scoped_lock lock(_inputMutex);
    size_t removedBytes = 0u;
    for (std::string& bytes : _inputQueue)
    {
        removedBytes += bytes.size();
        SecureZeroMemory(bytes.data(), bytes.size());
    }
    _inputQueue.clear();
    _inputQueuedBytes -= std::min(_inputQueuedBytes, removedBytes);
}

void Terminal::markUserInput(std::string_view bytes) noexcept
{
    if (bytes.empty())
    {
        return;
    }
    _lastInputTimestampNs.store(SteadyTimestampNs(), std::memory_order_release);
    {
        std::scoped_lock stateLock(_stateMutex);
        if (! _integrationTrusted)
        {
            return;
        }
        const bool submitted = bytes.find_first_of("\r\n\x03") != std::string_view::npos;
        const bool commandSubmitted = _commandSubmitted || submitted;
        const bool idleAtPrimaryPrompt = submitted ? false : _idleAtPrimaryPrompt;
        const bool hasPendingUserInput = submitted ? false : true;
        if (commandSubmitted == _commandSubmitted && idleAtPrimaryPrompt == _idleAtPrimaryPrompt &&
            hasPendingUserInput == _hasPendingUserInput)
        {
            return;
        }
        _commandSubmitted = commandSubmitted;
        _idleAtPrimaryPrompt = idleAtPrimaryPrompt;
        _hasPendingUserInput = hasPendingUserInput;
        _stateGeneration.fetch_add(1u, std::memory_order_acq_rel);
    }
}

bool Terminal::writeTextInput(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return true;
    }
    const std::string utf8 = Common::Strings::Utf8FromUtf16ReplacingInvalid(text);
    if (utf8.empty())
    {
        return false;
    }
    if (! writeInput(utf8))
    {
        return false;
    }
    markUserInput(utf8);
    return true;
}

HRESULT Terminal::ensureCloseWork() noexcept
{
    if (_closeWork)
    {
        return S_OK;
    }

    _closeWork.reset(CreateThreadpoolWork(&Terminal::CloseThreadpoolCallback, this, nullptr));
    if (! _closeWork)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    _closeModulePin = AcquireModuleReferenceFromAddress(&kTerminalModuleAnchor);
    if (! _closeModulePin)
    {
        _closeWork.reset();
        return E_FAIL;
    }
    return S_OK;
}

void Terminal::retireCommandSurfaceWorker() noexcept
{
    _commandSurfaceWorker.request_stop();
    if (! _commandSurfaceWorker.joinable())
    {
        return;
    }

    wil::unique_hmodule pin = AcquireModuleReferenceFromAddress(&kTerminalModuleAnchor);
    auto context = pin ? std::unique_ptr<CommandSurfaceJoinContext>(new (std::nothrow) CommandSurfaceJoinContext(
                             std::move(_commandSurfaceWorker), std::move(pin), this))
                       : nullptr;
    if (! context)
    {
        if (_commandSurfaceWorker.joinable())
        {
            _retiredCommandSurfaceWorkers.push_back(std::move(_commandSurfaceWorker));
        }
        return;
    }

    {
        std::scoped_lock lock(_commandSurfaceJoinMutex);
        ++_commandSurfaceJoinOutstanding;
    }
    AddRef();
    if (TrySubmitThreadpoolCallback(&Terminal::CommandSurfaceJoinCallback, context.get(), nullptr) == FALSE)
    {
        Release();
        {
            std::scoped_lock lock(_commandSurfaceJoinMutex);
            --_commandSurfaceJoinOutstanding;
        }
        _commandSurfaceJoinCv.notify_all();
        _retiredCommandSurfaceWorkers.push_back(std::move(context->worker));
        return;
    }
    static_cast<void>(context.release());
}

void CALLBACK Terminal::CommandSurfaceJoinCallback(PTP_CALLBACK_INSTANCE instance, void* context) noexcept
{
    std::unique_ptr<CommandSurfaceJoinContext> owned(static_cast<CommandSurfaceJoinContext*>(context));
    if (! owned)
    {
        return;
    }
    TransferModulePinToCallbackReturn(instance, owned->modulePin);
#if defined(ENABLE_TESTS)
    if (owned->terminal != nullptr)
    {
        owned->terminal->_debugCommandSurfaceJoinTid.store(GetCurrentThreadId(), std::memory_order_release);
    }
#endif
    if (owned->worker.joinable())
    {
        owned->worker.join();
    }
    if (owned->terminal != nullptr)
    {
        {
            std::scoped_lock lock(owned->terminal->_commandSurfaceJoinMutex);
            --owned->terminal->_commandSurfaceJoinOutstanding;
        }
        owned->terminal->_commandSurfaceJoinCv.notify_all();
        owned->terminal->Release();
    }
}

uint32_t Terminal::commandSurfaceJoinsOutstanding() noexcept
{
    std::scoped_lock lock(_commandSurfaceJoinMutex);
    return _commandSurfaceJoinOutstanding;
}

void Terminal::joinCommandSurfaceWorkers() noexcept
{
#if defined(ENABLE_TESTS)
    _debugCommandSurfaceJoinTid.store(GetCurrentThreadId(), std::memory_order_release);
#endif
    _commandSurfaceWorker.request_stop();
    if (_commandSurfaceWorker.joinable())
    {
        _commandSurfaceWorker.join();
    }
    for (std::jthread& worker : _retiredCommandSurfaceWorkers)
    {
        worker.request_stop();
        if (worker.joinable())
        {
            worker.join();
        }
    }
    _retiredCommandSurfaceWorkers.clear();
    {
        std::unique_lock lock(_commandSurfaceJoinMutex);
        _commandSurfaceJoinCv.wait(lock, [this]() noexcept { return _commandSurfaceJoinOutstanding == 0u; });
    }
}

void Terminal::prepareClose() noexcept
{
    flushShortcutRouteMetrics();
    setLifecycle(TerminalLifecycleState::Closing);
    ++_commandSurfaceGeneration;
    ++_commandSurfaceSnapshotGeneration;
    retireCommandSurfaceWorker();
    SecureZeroMemory(_commandSurfaceQuery.data(), _commandSurfaceQuery.size() * sizeof(wchar_t));
    if (_commandSurfaceSnapshot)
    {
        SecureZeroMemory(_commandSurfaceSnapshot->data(), _commandSurfaceSnapshot->size() * sizeof(wchar_t));
    }
    _commandSurfaceQuery.clear();
    _commandSurfaceSnapshot.reset();
    if (_commandSurfaceSessionHistorySnapshot)
    {
        ScrubStrings(*_commandSurfaceSessionHistorySnapshot);
    }
    _commandSurfaceSessionHistorySnapshot.reset();
    _commandSurfaceSnapshotReady = false;
    if (_persistedHistory)
    {
        ScrubStrings(*_persistedHistory);
    }
    _persistedHistory.reset();
    _persistedHistoryPath.clear();
    SecureZeroMemory(_commandSurfacePromptSnapshot.data(), _commandSurfacePromptSnapshot.size() * sizeof(wchar_t));
    _commandSurfacePromptSnapshot.clear();
    _commandSurfaceHistoryPath.clear();
    SecureZeroMemory(_commandSurfaceStatusOverride.data(), _commandSurfaceStatusOverride.size() * sizeof(wchar_t));
    _commandSurfaceStatusOverride.clear();
    ScrubFindRows(_commandSurfaceFindRows);
    ScrubSuggestionRows(_commandSurfaceSuggestionRows);
    _commandSurfaceKind = CommandSurfaceKind::None;
    _commandSurfaceInsertionPending = false;
    _commandSurfaceHistoryTrusted = false;
    _unsafePasteConfirmationVisible = false;
    _hyperlinkConfirmationVisible = false;
    _osc52ConfirmationVisible = false;
    _translatedCharacterSuppressionPending = false;
    _osc52PostPending.store(false, std::memory_order_release);
    SecureZeroMemory(_pendingUnsafePaste.data(), _pendingUnsafePaste.size());
    _pendingUnsafePaste.clear();
    _pendingHyperlink.clear();
    SecureZeroMemory(_pendingOsc52Text.data(), _pendingOsc52Text.size() * sizeof(wchar_t));
    _pendingOsc52Text.clear();
    releaseReportedMouseButtons();
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr && GetCapture() == hwnd)
    {
        ReleaseCapture();
    }
    stopSelection();
    _kittyPipeline.RequestStop();
    if (_accessibility)
    {
        _accessibility->Retire();
        _accessibility.reset();
    }
    clearQueuedInput();
    if (_window)
    {
        _window.reset();
    }
    discardDeviceResources();
    _dwriteFactory.reset();
    _d2dFactory.reset();
}

void CALLBACK Terminal::CloseThreadpoolCallback(PTP_CALLBACK_INSTANCE instance, void* context, PTP_WORK /*work*/) noexcept
{
    auto* self = static_cast<Terminal*>(context);
    if (self == nullptr)
    {
        return;
    }
    TransferModulePinToCallbackReturn(instance, self->_closeModulePin);
    self->finishClose();
    self->Release();
}

void Terminal::finishClose() noexcept
{
    joinCommandSurfaceWorkers();
    _kittyPipeline.Stop();
    if (_integrationThread.joinable())
    {
        _integrationThread.request_stop();
    }
    clearQueuedInput();

    {
        std::scoped_lock teardownLock(_sessionTeardownMutex);
        // Kill-on-close first stops every producer in the shell-owned tree. The
        // process handle is only a fallback for a partially initialized session.
        _processJob.reset();
        if (_process)
        {
            DWORD exitCode = 0u;
            if (GetExitCodeProcess(_process.get(), &exitCode) != FALSE && exitCode == STILL_ACTIVE)
            {
                static_cast<void>(TerminateProcess(_process.get(), ERROR_CANCELLED));
            }
        }
        {
            std::scoped_lock ioLock(_ioMutex);
            _ptyInputWrite.reset();
        }

        // Keep the output reader alive while ClosePseudoConsole drains the ConPTY
        // pipe on pre-26100 systems. Only after HPCON is gone may the worker use
        // cancellation as a final escape and join the reader.
        _pseudoConsole.reset();
    }
    if (_processWatcherThread.joinable())
    {
        _processWatcherThread.request_stop();
    }
    if (_readerThread.joinable())
    {
        _readerThread.request_stop();
        static_cast<void>(CancelSynchronousIo(_readerThread.native_handle()));
    }
    if (_integrationThread.joinable())
    {
        _integrationThread.join();
    }
    {
        std::scoped_lock stateLock(_stateMutex);
        SecureZeroMemory(_trustedPromptBuffer.data(), _trustedPromptBuffer.size() * sizeof(wchar_t));
        SecureZeroMemory(_pendingPromptExpectedBuffer.data(), _pendingPromptExpectedBuffer.size() * sizeof(wchar_t));
        SecureZeroMemory(_pendingPromptReplacement.data(), _pendingPromptReplacement.size() * sizeof(wchar_t));
        _trustedPromptBuffer.clear();
        _pendingPromptExpectedBuffer.clear();
        _pendingPromptReplacement.clear();
        _trustedHistoryPath.clear();
        ScrubStrings(_trustedCurrentSessionHistory);
        _trustedHistoryProviderAvailable = false;
        _integrationTrusted = false;
    }
    if (_readerThread.joinable())
    {
        _readerThread.join();
    }
    if (_processWatcherThread.joinable())
    {
        _processWatcherThread.join();
    }
    if (_process)
    {
        _process.reset();
    }
    releaseEngine();
    {
        std::scoped_lock lock(_terminalMutex);
        _diagnosticText.clear();
        _engineInitializationError.clear();
    }
    _kittyPlacements.clear();
    _kittyImages.clear();
    _kittyTargetStorageGeneration = 0u;
    _kittyCacheEpoch = 0u;
    _kittyCachedBytes = 0u;
    setLifecycle(TerminalLifecycleState::Exited);
}
