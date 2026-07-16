#pragma once

#include "FileSystemPathIdentity.h"
#include "PlugInterfaces/FileSystem.h"

#include <Windows.h>

#include <atomic>
#include <filesystem>
#include <span>
#include <string>
#include <vector>

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

// One changed plan row scheduled for execution. `currentSource` tracks the live
// provider path, which can hop through a unique temp name while breaking cycles;
// `originalSource` keeps the row identity for undo, success reporting, and
// per-item bookkeeping.
struct BatchRenameExecutionOp final
{
    std::filesystem::path currentSource;
    std::filesystem::path originalSource;
    std::wstring finalLeaf;
    std::filesystem::path tempPath;
    size_t depth              = 0u;
    bool isDirectory          = false;
    bool completed            = false;
    bool failed               = false;
    bool tempRestoreAttempted = false;
};

struct BatchRenameExecutionResult final
{
    HRESULT hr = S_OK;
    std::wstring detail;
    BatchRenameExecutionReport report;
    std::vector<std::filesystem::path> successfulSourcePaths;
    std::vector<std::filesystem::path> successfulTargetPaths;
    std::vector<ExecutedDirectoryMove> executedDirectoryMoves;
};

using BatchRenameExecutionProgressCallback = void (*)(void* context, uint64_t completedItems, uint64_t totalItems, bool forcePost) noexcept;

struct BatchRenameExecutionOptions final
{
    BatchRenameExecutionProgressCallback progressCallback = nullptr;
    void* progressContext                                 = nullptr;
};

[[nodiscard]] bool IsBatchRenameCancellationHRESULT(HRESULT hr) noexcept;
[[nodiscard]] size_t PathDepthKey(const std::filesystem::path& path) noexcept;
[[nodiscard]] std::filesystem::path JoinFolderAndLeaf(const FileSystemPathIdentity& pathIdentity,
                                                      const std::filesystem::path& folder,
                                                      std::wstring_view leaf) noexcept;
[[nodiscard]] std::filesystem::path ApplyExecutedDirectoryMoves(const FileSystemPathIdentity& pathIdentity,
                                                                std::filesystem::path path,
                                                                std::span<const ExecutedDirectoryMove> directoryMoves);
[[nodiscard]] BatchRenameExecutionResult RunBatchRenameExecutionEngine(std::atomic_bool& cancelRequested,
                                                                       IFileSystem& fileSystem,
                                                                       FileSystemPathIdentity pathIdentity,
                                                                       std::vector<BatchRenameExecutionOp> ops,
                                                                       BatchRenameExecutionOptions options = {}) noexcept;
