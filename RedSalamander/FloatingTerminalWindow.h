#pragma once

#include "AppTheme.h"
#include "PlugInterfaces/Terminal.h"
#include "SettingsStore.h"

#include <optional>
#include <string>
#include <string_view>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

struct CommandRuntimeState;

struct FloatingTerminalOpenRequest final
{
    std::wstring profileId;
    std::wstring providerId;
    std::wstring canonicalPath;
};

[[nodiscard]] HRESULT ShowFloatingTerminalWindow(HWND activationSource,
                                                 Common::Settings::Settings& settings,
                                                 const FloatingTerminalOpenRequest& request,
                                                 const AppTheme& theme) noexcept;
[[nodiscard]] HRESULT RestoreFloatingTerminalWindowAfterStartup(HWND activationSource, Common::Settings::Settings& settings, const AppTheme& theme) noexcept;
void PrepareFloatingTerminalWindowForAppShutdown() noexcept;
void UpdateFloatingTerminalWindowTheme(const AppTheme& theme) noexcept;
[[nodiscard]] HWND GetFloatingTerminalWindowHandle() noexcept;
[[nodiscard]] bool IsFloatingTerminalInputTarget(HWND targetWindow) noexcept;
[[nodiscard]] HRESULT RouteFloatingTerminalShortcut(
    HWND targetWindow, std::wstring_view commandId, const MSG& message, uint32_t modifiers, TerminalShortcutRoute& route) noexcept;
[[nodiscard]] bool ExecuteFloatingTerminalCommand(std::wstring_view commandId) noexcept;
[[nodiscard]] bool QueryFloatingTerminalCommandState(std::wstring_view commandId, CommandRuntimeState& state) noexcept;
[[nodiscard]] std::optional<FloatingTerminalOpenRequest> GetActiveFloatingTerminalRequest() noexcept;

#if defined(ENABLE_TESTS)
struct FloatingTerminalDebugSnapshot final
{
    HWND root                                   = nullptr;
    HWND selectedChild                          = nullptr;
    TerminalLifecycleState selectedLifecycle    = TerminalLifecycleState::Created;
    TerminalActivityTrust selectedActivityTrust = TerminalActivityTrust::Untrusted;
    uint64_t selectedSessionGeneration          = 0u;
    uint64_t pendingExitPayloadCount            = 0u;
    bool selectedFinalSnapshotComplete          = false;
    bool selectedIdleAtPrimaryPrompt            = false;
    size_t tabCount                             = 0u;
    size_t selectedIndex                        = 0u;
    std::vector<std::wstring> tabIds;
    std::vector<std::wstring> paths;
};

[[nodiscard]] bool DebugGetFloatingTerminalSnapshot(FloatingTerminalDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugCloseFloatingTerminalTab(size_t index) noexcept;
[[nodiscard]] bool DebugReorderFloatingTerminalTab(size_t fromIndex, size_t toIndex) noexcept;
[[nodiscard]] HRESULT DebugTerminateFloatingTerminalRootProcess(uint32_t exitCode) noexcept;
void DebugGetFloatingTerminalRootExitTiming(uint64_t& generation, uint64_t& durationUs) noexcept;
#endif
