#pragma once

#include <Windows.h>

#include <algorithm>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <exception>
#include <filesystem>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <utility>

namespace RedSalamander::TestSupport
{
inline constexpr std::wstring_view kTestRootEnvironmentVariable{L"REDSALAMANDER_TEST_ROOT"};
inline constexpr std::wstring_view kTestRunIdEnvironmentVariable{L"REDSALAMANDER_TEST_RUN_ID"};
inline constexpr std::wstring_view kRunsDirectoryName{L"runs"};
inline constexpr std::wstring_view kScratchDirectoryName{L"scratch"};
inline constexpr std::wstring_view kArtifactsDirectoryName{L"artifacts"};

enum class TestDirectoryKind : uint8_t
{
    Scratch,
    Artifacts,
};

struct TestDirectoryOptions final
{
    std::wstring_view harnessSegment;
    std::wstring_view leafSegment;
    std::wstring_view fallbackRunIdPrefix;
    std::wstring_view emptyLeafFallback = L"case";
    TestDirectoryKind kind              = TestDirectoryKind::Scratch;
    bool includeLeafSegment             = true;
    bool cleanExisting                  = true;
};

[[nodiscard]] inline bool IsSafeTestSandboxSegmentCharacter(const wchar_t ch) noexcept
{
    return (ch >= L'0' && ch <= L'9') || (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z') || ch == L'.' || ch == L'-' || ch == L'_';
}

[[nodiscard]] inline std::wstring SanitizeTestSandboxSegment(std::wstring_view text, std::wstring_view emptyFallback = L"case")
{
    std::wstring result;
    result.reserve(text.empty() ? emptyFallback.size() : text.size());
    for (const wchar_t ch : text)
    {
        result.push_back(IsSafeTestSandboxSegmentCharacter(ch) ? ch : L'_');
    }

    if (result.empty())
    {
        result.assign(emptyFallback);
    }
    return result;
}

[[nodiscard]] inline bool IsSafeTestRunId(std::wstring_view text) noexcept
{
    if (text.empty() || text.size() > 160u)
    {
        return false;
    }

    for (const wchar_t ch : text)
    {
        if (! IsSafeTestSandboxSegmentCharacter(ch))
        {
            return false;
        }
    }
    return true;
}

struct EnvironmentValue final
{
    std::optional<std::wstring> value;
    DWORD error = ERROR_SUCCESS;
};

[[nodiscard]] inline EnvironmentValue ReadEnvironmentValue(std::wstring_view name)
{
    const std::wstring variableName(name);
    SetLastError(ERROR_SUCCESS);
    const DWORD required = GetEnvironmentVariableW(variableName.c_str(), nullptr, 0u);
    if (required == 0u)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_SUCCESS)
        {
            return {.value = std::wstring{}};
        }
        if (error == ERROR_ENVVAR_NOT_FOUND)
        {
            return {};
        }
        return {.error = error};
    }

    std::wstring value(required, L'\0');
    SetLastError(ERROR_SUCCESS);
    const DWORD written = GetEnvironmentVariableW(variableName.c_str(), value.data(), required);
    if (written == 0u || written >= required)
    {
        DWORD error = GetLastError();
        if (error == ERROR_SUCCESS)
        {
            error = ERROR_INVALID_DATA;
        }
        return {.error = error};
    }

    value.resize(written);
    return {.value = std::move(value)};
}

[[nodiscard]] inline std::wstring GetEnvironmentString(std::wstring_view name)
{
    EnvironmentValue result = ReadEnvironmentValue(name);
    return result.value.value_or(std::wstring{});
}

class ScopedEnvironmentVariable final
{
public:
    ScopedEnvironmentVariable(std::wstring_view name, std::optional<std::wstring_view> value) : _name(name)
    {
        EnvironmentValue original = ReadEnvironmentValue(_name);
        if (original.error != ERROR_SUCCESS)
        {
            throw std::system_error(static_cast<int>(original.error), std::system_category(), "GetEnvironmentVariableW");
        }
        _original = std::move(original.value);

        const std::wstring replacement = value ? std::wstring(*value) : std::wstring{};
        if (SetEnvironmentVariableW(_name.c_str(), value ? replacement.c_str() : nullptr) == FALSE)
        {
            throw std::system_error(static_cast<int>(GetLastError()), std::system_category(), "SetEnvironmentVariableW");
        }
        _active = true;
    }

    ScopedEnvironmentVariable(const ScopedEnvironmentVariable&)            = delete;
    ScopedEnvironmentVariable& operator=(const ScopedEnvironmentVariable&) = delete;
    ScopedEnvironmentVariable(ScopedEnvironmentVariable&&)                 = delete;
    ScopedEnvironmentVariable& operator=(ScopedEnvironmentVariable&&)      = delete;

    ~ScopedEnvironmentVariable() noexcept
    {
        if (_active)
        {
            static_cast<void>(SetEnvironmentVariableW(_name.c_str(), _original ? _original->c_str() : nullptr));
        }
    }

private:
    std::wstring _name;
    std::optional<std::wstring> _original;
    bool _active = false;
};

[[nodiscard]] inline std::filesystem::path ResolveTestSandboxBase(std::error_code& ec) noexcept
{
    ec.clear();
    try
    {
        const std::wstring configuredRoot = GetEnvironmentString(kTestRootEnvironmentVariable);
        if (! configuredRoot.empty())
        {
            return std::filesystem::path(configuredRoot).lexically_normal();
        }

        std::filesystem::path current = std::filesystem::current_path(ec);
        if (ec)
        {
            return {};
        }

        current                                  = current.lexically_normal();
        const std::filesystem::path platformRoot = current.parent_path();
        const std::filesystem::path buildRoot    = platformRoot.parent_path();
        if ((current.filename() == L"Debug" || current.filename() == L"Release" || current.filename() == L"ASan Debug") && buildRoot.filename() == L".build")
        {
            return (buildRoot / L"TestSandbox").lexically_normal();
        }

        return (current / L".build" / L"TestSandbox").lexically_normal();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // This is test infrastructure: convert an unexpected path/runtime failure into the existing error-code contract.
        ec = std::make_error_code(std::errc::io_error);
        return {};
    }
}

[[nodiscard]] inline std::wstring ResolveTestRunId(std::wstring_view fallbackPrefix)
{
    const std::wstring configuredRunId = GetEnvironmentString(kTestRunIdEnvironmentVariable);
    if (IsSafeTestRunId(configuredRunId))
    {
        return configuredRunId;
    }

    std::wstring runId(fallbackPrefix);
    runId.append(L"-");
    runId.append(std::to_wstring(GetCurrentProcessId()));
    runId.append(L"-");
    runId.append(std::to_wstring(GetTickCount64()));
    return runId;
}

[[nodiscard]] inline std::filesystem::path AcquireTestDirectory(const TestDirectoryOptions& options, std::error_code& ec) noexcept
{
    ec.clear();
    try
    {
        const std::filesystem::path sandboxBase = ResolveTestSandboxBase(ec);
        if (ec || sandboxBase.empty())
        {
            return {};
        }

        const std::wstring runId         = ResolveTestRunId(options.fallbackRunIdPrefix);
        const std::wstring_view areaName = options.kind == TestDirectoryKind::Scratch ? kScratchDirectoryName : kArtifactsDirectoryName;
        std::filesystem::path directory =
            sandboxBase / std::wstring(kRunsDirectoryName) / runId / std::wstring(areaName) / std::wstring(options.harnessSegment);
        if (options.includeLeafSegment)
        {
            directory /= SanitizeTestSandboxSegment(options.leafSegment, options.emptyLeafFallback);
        }
        directory = directory.lexically_normal();

        if (options.cleanExisting)
        {
            std::filesystem::remove_all(directory, ec);
            if (ec)
            {
                return {};
            }
        }

        std::filesystem::create_directories(directory, ec);
        if (ec)
        {
            return {};
        }
        return directory;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Keep standalone test harnesses on their existing error-code failure path.
        ec = std::make_error_code(std::errc::io_error);
        return {};
    }
}

struct MessagePumpWaitOptions final
{
    std::chrono::milliseconds timeout{};
    std::chrono::milliseconds pollInterval{20};
    size_t maxMessagesPerPump = 1024u;
    std::wstring_view operationName{L"test condition"};
};

struct MessagePumpWaitResult final
{
    bool conditionMet             = false;
    size_t dispatchedMessageCount = 0u;
    std::chrono::milliseconds elapsed{};
    std::wstring timeoutDiagnostic;
};

[[nodiscard]] inline size_t PumpPendingMessages(size_t maxMessageCount = 1024u) noexcept
{
    size_t dispatchedCount = 0u;
    MSG msg{};
    while (dispatchedCount < maxMessageCount && PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
        ++dispatchedCount;
    }
    return dispatchedCount;
}

template <typename Predicate> [[nodiscard]] inline MessagePumpWaitResult PumpMessagesUntil(Predicate&& predicate, const MessagePumpWaitOptions& options)
{
    MessagePumpWaitResult result{};
    const auto startedAt = std::chrono::steady_clock::now();
    const auto deadline  = startedAt + options.timeout;

    while (std::chrono::steady_clock::now() < deadline)
    {
        result.dispatchedMessageCount += PumpPendingMessages(options.maxMessagesPerPump);
        if (predicate())
        {
            result.conditionMet = true;
            break;
        }

        const auto remaining = std::chrono::duration_cast<std::chrono::milliseconds>(deadline - std::chrono::steady_clock::now());
        if (remaining > std::chrono::milliseconds::zero())
        {
            std::this_thread::sleep_for((std::min)(options.pollInterval, remaining));
        }
    }

    if (! result.conditionMet)
    {
        result.dispatchedMessageCount += PumpPendingMessages(options.maxMessagesPerPump);
        result.conditionMet = predicate();
    }

    result.elapsed = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - startedAt);
    if (! result.conditionMet)
    {
        result.timeoutDiagnostic.assign(options.operationName);
        result.timeoutDiagnostic.append(L" timed out after ");
        result.timeoutDiagnostic.append(std::to_wstring(result.elapsed.count()));
        result.timeoutDiagnostic.append(L" ms (budget ");
        result.timeoutDiagnostic.append(std::to_wstring(options.timeout.count()));
        result.timeoutDiagnostic.append(L" ms, dispatched ");
        result.timeoutDiagnostic.append(std::to_wstring(result.dispatchedMessageCount));
        result.timeoutDiagnostic.append(L" messages)");
    }
    return result;
}

template <typename Snapshot, typename CaptureSnapshot, typename Predicate>
[[nodiscard]] inline bool WaitForSnapshot(CaptureSnapshot&& captureSnapshot,
                                          Predicate&& predicate,
                                          const MessagePumpWaitOptions& options,
                                          Snapshot* outSnapshot           = nullptr,
                                          std::wstring* timeoutDiagnostic = nullptr)
{
    Snapshot snapshot{};
    Snapshot lastSnapshot{};
    bool sawSnapshot                 = false;
    MessagePumpWaitResult waitResult = PumpMessagesUntil(
        [&]() noexcept
    {
        if (! captureSnapshot(snapshot))
        {
            return false;
        }

        lastSnapshot = snapshot;
        sawSnapshot  = true;
        if (! predicate(snapshot))
        {
            return false;
        }

        if (outSnapshot)
        {
            *outSnapshot = snapshot;
        }
        return true;
    },
        options);

    if (! waitResult.conditionMet && outSnapshot && sawSnapshot)
    {
        *outSnapshot = lastSnapshot;
    }
    if (timeoutDiagnostic)
    {
        *timeoutDiagnostic = std::move(waitResult.timeoutDiagnostic);
    }
    return waitResult.conditionMet;
}
} // namespace RedSalamander::TestSupport
