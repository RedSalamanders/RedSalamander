#pragma once

#include <cstddef>
#include <filesystem>
#include <cstdint>
#include <span>
#include <string>
#include <string_view>

#include <Windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL: deleted copy/move operators
#include <wil/resource.h>
#pragma warning(pop)

#include "SettingsStore.h"

namespace FileActionLauncher
{
#if defined(ENABLE_TESTS)
namespace Testing
{
struct SelectedPathsFileObservation final
{
    std::filesystem::path collidedPath;
    uint64_t recordCount             = 0u;
    uint64_t pathBytes               = 0u;
    uint64_t writeCallCount          = 0u;
    uint64_t peakAdditionalBytes     = 0u;
    uint64_t selectionCopyBytes      = 0u;
    size_t createAttempts            = 0u;
    bool usedAggregatePayload        = false;
};

struct SelectedPathsFileOptions final
{
    std::filesystem::path rootOverride;
    size_t injectedCreateCollisions = 0u;
    size_t maximumCreateAttempts    = 32u;
    SelectedPathsFileObservation* observation = nullptr;
};

struct SelectedPathsRecoveryObservation final
{
    uint64_t inspected = 0u;
    uint64_t deleted   = 0u;
    uint64_t skipped   = 0u;
    uint64_t errors    = 0u;
    uint64_t elapsedUs = 0u;
    bool stoppedByEntryBound  = false;
    bool stoppedByDeleteBound = false;
    bool stoppedByTimeBound   = false;
    bool resetMissingCursor   = false;
};
} // namespace Testing
#endif

struct MacroContext
{
    std::filesystem::path itemPath;
    std::filesystem::path currentDirectory;
    std::filesystem::path oppositePanePath;
    std::filesystem::path selectedPathsFile;
    std::span<const std::filesystem::path> selectedPaths;
    std::wstring computerName;
#if defined(ENABLE_TESTS)
    const Testing::SelectedPathsFileOptions* selectedPathsFileOptions = nullptr;
#endif
};

struct ExternalActionPathReferences final
{
    bool itemPath         = false;
    bool currentDirectory = false;
    bool oppositePanePath = false;
    bool selectedPathsFile = false;
};

struct LaunchPlan
{
    LaunchPlan()                              = default;
    LaunchPlan(const LaunchPlan&)             = delete;
    LaunchPlan& operator=(const LaunchPlan&)  = delete;
    LaunchPlan(LaunchPlan&&) noexcept         = default;
    LaunchPlan& operator=(LaunchPlan&&) noexcept = default;

    std::wstring executablePath;
    std::wstring arguments;
    std::wstring workingDirectory;
    class SelectedPathsFileLease
    {
    public:
        SelectedPathsFileLease() noexcept = default;
        SelectedPathsFileLease(std::filesystem::path path,
                               std::filesystem::path rootPath,
                               const FILE_ID_INFO& fileIdentity,
                               const FILE_ID_INFO& rootIdentity) noexcept;
        ~SelectedPathsFileLease() noexcept;

        SelectedPathsFileLease(const SelectedPathsFileLease&)            = delete;
        SelectedPathsFileLease& operator=(const SelectedPathsFileLease&) = delete;
        SelectedPathsFileLease(SelectedPathsFileLease&& other) noexcept;
        SelectedPathsFileLease& operator=(SelectedPathsFileLease&& other) noexcept;

        [[nodiscard]] const std::filesystem::path& Path() const noexcept;
        [[nodiscard]] bool HasValue() const noexcept;
        void Reset() noexcept;

    private:
        std::filesystem::path _path;
        std::filesystem::path _rootPath;
        FILE_ID_INFO _fileIdentity{};
        FILE_ID_INFO _rootIdentity{};
        bool _hasIdentity = false;
    } selectedPathsFileLease;
};

struct LaunchOptions
{
    HWND ownerWindow          = nullptr;
    int showCommand           = SW_SHOWNORMAL;
    bool waitForExit          = false;
    DWORD waitTimeoutMs       = INFINITE;
    bool captureProcessHandle = false;
    DWORD selectedPathsMaximumRetentionMs = 10u * 60u * 1000u;
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
[[nodiscard]] HRESULT GetExternalActionPathReferences(const Common::Settings::FileActionDefinition& action,
                                                      ExternalActionPathReferences& references) noexcept;
[[nodiscard]] HRESULT ExternalActionUsesSelectedPathsFile(const Common::Settings::FileActionDefinition& action,
                                                          bool& usesSelectedPathsFile) noexcept;
[[nodiscard]] HRESULT BuildExternalLaunchPlan(const Common::Settings::FileActionDefinition& action, const MacroContext& context, LaunchPlan& out) noexcept;
[[nodiscard]] HRESULT LaunchExternalPlan(LaunchPlan plan, const LaunchOptions& options = {}, LaunchResult* result = nullptr) noexcept;

#if defined(ENABLE_TESTS)
namespace Testing
{
enum class ExternalLaunchFault : uint8_t
{
    None,
    DeferredCleanupAllocation,
    DeferredCleanupWaitCreation,
    ShellExecute,
    NullProcessHandle,
    WaitFailure,
    ExitCodeQuery,
};

void SetNextExternalLaunchFault(ExternalLaunchFault fault) noexcept;
[[nodiscard]] HRESULT RunSelectedPathsRecoveryPassForTest(const std::filesystem::path& root,
                                                          uint64_t minimumAgeMs,
                                                          size_t maximumEntries,
                                                          size_t maximumDeletes,
                                                          uint64_t maximumElapsedMs,
                                                          std::wstring& cursor,
                                                          SelectedPathsRecoveryObservation& observation) noexcept;
} // namespace Testing
#endif
} // namespace FileActionLauncher
