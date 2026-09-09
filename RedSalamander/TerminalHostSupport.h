#pragma once

#include "AppTheme.h"
#include "PlugInterfaces/Terminal.h"

#include <filesystem>
#include <string>
#include <string_view>

namespace TerminalHostSupport
{
struct OwnedTerminalLocation final
{
    TerminalLocationKind kind = TerminalLocationKind::Unsupported;
    std::wstring windowsPath;
    std::wstring wslDistribution;
    std::wstring wslAbsolutePath;
    std::wstring pluginShortId;
    std::wstring pluginBackingWindowsPath;

    [[nodiscard]] TerminalLogicalLocation View() const noexcept;
    [[nodiscard]] std::filesystem::path IdentityPath() const;
};

[[nodiscard]] TerminalUtf16Span TerminalSpan(std::wstring_view text) noexcept;
[[nodiscard]] OwnedTerminalLocation MakeTerminalLocation(const std::filesystem::path& path);
[[nodiscard]] TerminalTheme BuildTerminalTheme(const AppTheme& appTheme, uint32_t dpi) noexcept;
#if defined(ENABLE_TESTS)
[[nodiscard]] HRESULT DebugTerminateRootProcess(ITerminal* terminal, uint32_t exitCode) noexcept;
[[nodiscard]] HRESULT DebugGetScreenText(ITerminal* terminal, std::wstring& text) noexcept;
[[nodiscard]] HRESULT DebugRunCommandExperiencePerfSelfTests(
    unsigned int* passedTests,
    unsigned int* failedTests) noexcept;
#endif
} // namespace TerminalHostSupport
