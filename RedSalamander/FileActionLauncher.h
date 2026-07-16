#pragma once

#include <filesystem>
#include <string>
#include <string_view>
#include <vector>

#include <Windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL: deleted copy/move operators
#include <wil/resource.h>
#pragma warning(pop)

#include "SettingsStore.h"

namespace FileActionLauncher
{
struct MacroContext
{
    std::filesystem::path itemPath;
    std::filesystem::path currentDirectory;
    std::filesystem::path oppositePanePath;
    std::filesystem::path selectedPathsFile;
    std::vector<std::filesystem::path> selectedPaths;
    std::wstring computerName;
};

struct LaunchPlan
{
    std::wstring executablePath;
    std::wstring arguments;
    std::wstring workingDirectory;
    std::vector<std::filesystem::path> cleanupFilesAfterExit;
};

struct LaunchOptions
{
    HWND ownerWindow          = nullptr;
    int showCommand           = SW_SHOWNORMAL;
    bool waitForExit          = false;
    DWORD waitTimeoutMs       = INFINITE;
    bool captureProcessHandle = false;
};

struct LaunchResult
{
    LaunchResult()                                   = default;
    LaunchResult(const LaunchResult&)                = delete;
    LaunchResult& operator=(const LaunchResult&)     = delete;
    LaunchResult(LaunchResult&&) noexcept            = default;
    LaunchResult& operator=(LaunchResult&&) noexcept = default;

    bool exitCodeAvailable = false;
    DWORD exitCode         = 0;
    DWORD processId        = 0;
    wil::unique_handle processHandle;
};

[[nodiscard]] HRESULT ExpandMacros(std::wstring_view templateText, const MacroContext& context, std::wstring& out) noexcept;
[[nodiscard]] bool TemplateContainsSupportedMacro(std::wstring_view templateText) noexcept;
[[nodiscard]] HRESULT BuildExternalLaunchPlan(const Common::Settings::FileActionDefinition& action, const MacroContext& context, LaunchPlan& out) noexcept;
[[nodiscard]] HRESULT LaunchExternalPlan(const LaunchPlan& plan, const LaunchOptions& options = {}, LaunchResult* result = nullptr) noexcept;
} // namespace FileActionLauncher
