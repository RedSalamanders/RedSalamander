#pragma once

#include "CommandRegistry.h"
#include "PlugInterfaces/Terminal.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include <cstdint>
#include <cstring>
#include <string_view>

struct CommandRuntimeState final
{
    bool enabled                 = true;
    bool checked                 = false;
    bool terminalIdentityPresent = false;
    TerminalInstanceId terminalInstanceId{};
    uint64_t terminalSessionGeneration = 0u;
};

[[nodiscard]] inline bool CommandRuntimeIdentityMatches(const CommandRuntimeState& captured, const CommandRuntimeState& current) noexcept
{
    return captured.terminalIdentityPresent == current.terminalIdentityPresent &&
           (! captured.terminalIdentityPresent ||
            (captured.terminalSessionGeneration == current.terminalSessionGeneration &&
             std::memcmp(captured.terminalInstanceId.bytes, current.terminalInstanceId.bytes, sizeof(captured.terminalInstanceId.bytes)) == 0));
}

// Resolves the state source declared by CommandRegistry against the exact window
// that owned the invocation. Palette activation calls this again before dispatch,
// so a replaced Terminal instance cannot inherit an older row's authority.
[[nodiscard]] bool ResolveCommandRuntimeStateFromWindow(
    HWND ownerWindow, HWND invocationOrigin, std::wstring_view commandId, CommandStateSource stateSource, CommandRuntimeState& state) noexcept;
