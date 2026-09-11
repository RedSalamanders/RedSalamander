#pragma once

#include "FileSystemPathIdentity.h"
#include "PlugInterfaces/FileSystem.h"

#include <atomic>
#include <filesystem>
#include <span>
#include <string>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

struct BatchRenameUndoEntry final
{
    std::filesystem::path currentPath;
    std::wstring restoreName;
    std::filesystem::path originalPath;
};

struct BatchRenameExecutionReport final
{
    size_t totalRows     = 0u;
    size_t completedRows = 0u;
    size_t skippedRows   = 0u;
    size_t failedRows    = 0u;
    HRESULT firstFailure = S_OK;
    bool canceled        = false;
    std::wstring firstFailureText;
    std::vector<BatchRenameUndoEntry> undoEntries;
};

struct ExecutedDirectoryMove final
{
    std::filesystem::path sourcePath;
    std::filesystem::path targetPath;
};

// One changed plan row scheduled for execution. Cycles are rejected before mutation, so the
// admitted provider path never changes while the row waits for its layer.
struct BatchRenameExecutionOp final
{
    std::filesystem::path originalSource;
    std::wstring finalLeaf;
    std::filesystem::path providerFinalPath;
    std::wstring providerParentKey;
    std::wstring providerSourceCollisionKey;
    std::wstring providerFinalCollisionKey;
    size_t depth     = 0u;
    bool isDirectory = false;
    bool completed   = false;
    bool skipped     = false;
    bool failed      = false;
    HRESULT status   = E_PENDING;
};

struct BatchRenameExecutionLayer final
{
    std::vector<size_t> operationIndices;
};

// Immutable dependency result shared by preview, worker admission, plan validation,
// and execution. A successful schedule covers every operation exactly once. A cycle
// result has no executable layers and names only the exact cycle members.
struct BatchRenameExecutionSchedule final
{
    std::vector<BatchRenameExecutionLayer> layers;
    std::vector<size_t> cycleOperationIndices;
};

struct BatchRenameExecutionItemResult final
{
    size_t operationIndex = 0u;
    HRESULT status        = E_PENDING;
    bool completed        = false;
    bool skipped          = false;
};

struct BatchRenameExecutionResult final
{
    HRESULT hr = S_OK;
    std::wstring detail;
    BatchRenameExecutionReport report;
    std::vector<std::filesystem::path> successfulSourcePaths;
    std::vector<std::filesystem::path> successfulTargetPaths;
    std::vector<ExecutedDirectoryMove> executedDirectoryMoves;
    std::vector<BatchRenameExecutionItemResult> itemResults;
};

using BatchRenameExecutionProgressCallback = void (*)(void* context, uint64_t completedItems, uint64_t totalItems, bool forcePost) noexcept;

// The dependency scheduler never calls a provider mutation API. Its owner supplies the one
// identity-bound mutation boundary; every scheduled row is a final rename.
using BatchRenameMutationCallback = HRESULT (*)(void* context,
                                                size_t operationIndex,
                                                const std::filesystem::path& sourcePath,
                                                const std::filesystem::path& destinationPath) noexcept;

struct BatchRenameExecutionOptions final
{
    BatchRenameExecutionProgressCallback progressCallback = nullptr;
    void* progressContext                                 = nullptr;
    BatchRenameMutationCallback mutationCallback           = nullptr;
    void* mutationContext                                  = nullptr;
    const BatchRenameExecutionSchedule* schedule            = nullptr;
};

[[nodiscard]] bool IsBatchRenameCancellationHRESULT(HRESULT hr) noexcept;
[[nodiscard]] size_t PathDepthKey(const std::filesystem::path& path) noexcept;
[[nodiscard]] std::filesystem::path JoinFolderAndLeaf(const FileSystemPathIdentity& pathIdentity,
                                                      const std::filesystem::path& folder,
                                                      std::wstring_view leaf) noexcept;
[[nodiscard]] std::filesystem::path ApplyExecutedDirectoryMoves(const FileSystemPathIdentity& pathIdentity,
                                                                std::filesystem::path path,
                                                                std::span<const ExecutedDirectoryMove> directoryMoves);
[[nodiscard]] HRESULT BuildBatchRenameExecutionSchedule(const FileSystemPathIdentity& pathIdentity,
                                                        std::span<const BatchRenameExecutionOp> operations,
                                                        BatchRenameExecutionSchedule& out) noexcept;
[[nodiscard]] BatchRenameExecutionResult RunBatchRenameExecutionEngine(std::atomic_bool& cancelRequested,
                                                                       FileSystemPathIdentity pathIdentity,
                                                                       std::vector<BatchRenameExecutionOp> ops,
                                                                       BatchRenameExecutionOptions options = {}) noexcept;
