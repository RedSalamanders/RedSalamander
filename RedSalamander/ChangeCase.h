#pragma once

#include "BatchRenameExecutionEngine.h"
#include "framework.h"

#include <cstdint>
#include <filesystem>
#include <span>
#include <stop_token>
#include <string>
#include <string_view>
#include <vector>

struct IFileSystem;

namespace ChangeCase
{
enum class CaseStyle : uint8_t
{
    Lower,
    Upper,
    PartiallyMixed, // name in mixed, extension in lower (when applicable)
    Mixed,
};

enum class ChangeTarget : uint8_t
{
    WholeFilename,
    OnlyName,
    OnlyExtension,
};

struct Options
{
    CaseStyle style     = CaseStyle::Lower;
    ChangeTarget target = ChangeTarget::WholeFilename;
    bool includeSubdirs = false;
};

[[nodiscard]] std::wstring TransformLeafName(std::wstring_view leafName, const Options& options) noexcept;

#ifdef ENABLE_TESTS
[[nodiscard]] HRESULT DebugClassifyDirectoryReadResult(HRESULT readHr, bool hasInformation, bool provenDirectory) noexcept;
#endif

struct ProgressUpdate final
{
    enum class Phase : uint8_t
    {
        Enumerating,
        Renaming,
    };

    Phase phase = Phase::Enumerating;
    std::filesystem::path currentPath;

    uint64_t scannedFolders = 0;
    uint64_t scannedEntries = 0;

    uint64_t plannedRenames   = 0;
    uint64_t completedRenames = 0;
};

using ProgressCallback = void (*)(const ProgressUpdate& update, void* cookie) noexcept;

// Performs provider discovery and namespace validation only. It never mutates. The returned
// operations are immutable admission input for the central RenamePlan executor.
[[nodiscard]] HRESULT BuildRenameOperations(IFileSystem& fileSystem,
                                            std::wstring_view pluginId,
                                            const std::vector<std::filesystem::path>& inputPaths,
                                            const Options& options,
                                            std::vector<BatchRenameExecutionOp>& operationsOut,
                                            std::stop_token stopToken = {},
                                            ProgressCallback progress = nullptr,
                                            void* progressCookie      = nullptr) noexcept;

#ifdef ENABLE_TESTS
struct MutationGuardCallbacks final
{
    // Called once after iterative discovery/planning and before the first rename. The span contains
    // exactly the source paths that the operation plans to mutate.
    HRESULT (*prepare)(std::span<const std::filesystem::path> paths, void* cookie) noexcept = nullptr;

    // Called immediately before every provider rename batch. A failure prevents that batch and all
    // later batches from mutating.
    HRESULT (*revalidate)(std::span<const std::filesystem::path> paths, void* cookie) noexcept = nullptr;
    void* cookie                                                                               = nullptr;
};

// Test/compatibility adapter for direct engine characterization. Production command dispatch
// submits BuildRenameOperations() output to the central RenamePlan executor.
// Applies the requested case transformation to the given paths.
// Notes:
// - includeSubdirs uses IFileSystem::ReadDirectoryInfo (non-recursive traversal).
// - Renames are batched via IFileSystem::RenameItems.
// - stopToken allows cooperative cancellation.
[[nodiscard]] HRESULT DebugApplyToPathsForTests(IFileSystem& fileSystem,
                                                std::wstring_view pluginId,
                                                const std::vector<std::filesystem::path>& inputPaths,
                                                const Options& options,
                                                std::stop_token stopToken                   = {},
                                                ProgressCallback progress                   = nullptr,
                                                void* progressCookie                        = nullptr,
                                                const MutationGuardCallbacks* mutationGuard = nullptr) noexcept;
#endif
} // namespace ChangeCase
