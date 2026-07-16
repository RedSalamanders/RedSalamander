#include "Framework.h"

#include "BatchRenameExecutionEngine.h"

#include "FileSystemRenameBatch.h"
#include "Helpers.h"

#include <algorithm>
#include <chrono>
#include <format>
#include <limits>
#include <mutex>
#include <optional>
#include <utility>

namespace
{
constexpr ULONGLONG kBatchRenameProgressThrottleMs = 100ull;

struct BatchRenameItemOutcome final
{
    unsigned long itemIndex = ULONG_MAX;
    std::wstring sourcePath;
    HRESULT status = S_OK;
};

class BatchRenameExecutionCallback final : public IFileSystemCallback
{
public:
    BatchRenameExecutionCallback(const std::atomic_bool& cancelRequested, BatchRenameExecutionOptions options, const size_t overallTotal) noexcept
        : _cancelRequested(cancelRequested),
          _options(options),
          _overallTotal(overallTotal)
    {
    }

    BatchRenameExecutionCallback(const BatchRenameExecutionCallback&)            = delete;
    BatchRenameExecutionCallback& operator=(const BatchRenameExecutionCallback&) = delete;
    BatchRenameExecutionCallback(BatchRenameExecutionCallback&&)                 = delete;
    BatchRenameExecutionCallback& operator=(BatchRenameExecutionCallback&&)      = delete;

    void SetProgressBase(const size_t processedOps) noexcept
    {
        _progressBase.store(processedOps, std::memory_order_release);
    }

    void NotifyOverallProgress(const size_t processedOps) noexcept
    {
        PostProgress(static_cast<uint64_t>(processedOps), processedOps >= _overallTotal);
    }

    void ResetItemOutcomes() noexcept
    {
        std::lock_guard lock(_outcomesMutex);
        _itemOutcomes.clear();
    }

    [[nodiscard]] std::vector<BatchRenameItemOutcome> TakeItemOutcomes() noexcept
    {
        std::lock_guard lock(_outcomesMutex);
        return std::exchange(_itemOutcomes, {});
    }

    HRESULT STDMETHODCALLTYPE FileSystemProgress([[maybe_unused]] FileSystemOperation operationType,
                                                 unsigned long totalItems,
                                                 unsigned long completedItems,
                                                 [[maybe_unused]] uint64_t totalBytes,
                                                 [[maybe_unused]] uint64_t completedBytes,
                                                 [[maybe_unused]] const wchar_t* currentSourcePath,
                                                 [[maybe_unused]] const wchar_t* currentDestinationPath,
                                                 [[maybe_unused]] uint64_t currentItemTotalBytes,
                                                 [[maybe_unused]] uint64_t currentItemCompletedBytes,
                                                 [[maybe_unused]] FileSystemOptions* options,
                                                 [[maybe_unused]] uint64_t progressStreamId,
                                                 [[maybe_unused]] void* cookie) noexcept override
    {
        const uint64_t overallCompleted =
            std::min(static_cast<uint64_t>(_progressBase.load(std::memory_order_acquire)) + completedItems, static_cast<uint64_t>(_overallTotal));
        PostProgress(overallCompleted, totalItems != 0u && completedItems >= totalItems);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemItemCompleted([[maybe_unused]] FileSystemOperation operationType,
                                                      unsigned long itemIndex,
                                                      const wchar_t* sourcePath,
                                                      [[maybe_unused]] const wchar_t* destinationPath,
                                                      HRESULT status,
                                                      [[maybe_unused]] FileSystemOptions* options,
                                                      [[maybe_unused]] void* cookie) noexcept override
    {
        if (! sourcePath)
        {
            return S_OK;
        }

        std::lock_guard lock(_outcomesMutex);
        _itemOutcomes.push_back(BatchRenameItemOutcome{.itemIndex = itemIndex, .sourcePath = sourcePath, .status = status});
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, [[maybe_unused]] void* cookie) noexcept override
    {
        if (! pCancel)
        {
            return E_POINTER;
        }

        *pCancel = _cancelRequested.load(std::memory_order_acquire) ? TRUE : FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemIssue([[maybe_unused]] FileSystemOperation operationType,
                                              [[maybe_unused]] const wchar_t* sourcePath,
                                              [[maybe_unused]] const wchar_t* destinationPath,
                                              [[maybe_unused]] HRESULT status,
                                              FileSystemIssueAction* action,
                                              [[maybe_unused]] FileSystemOptions* options,
                                              [[maybe_unused]] void* cookie) noexcept override
    {
        if (! action)
        {
            return E_POINTER;
        }

        *action = FileSystemIssueAction::Cancel;
        return S_OK;
    }

private:
    void PostProgress(const uint64_t overallCompleted, const bool forcePost) noexcept
    {
        if (! _options.progressCallback)
        {
            return;
        }
        const uint64_t clampedCompleted = std::min(overallCompleted, static_cast<uint64_t>(_overallTotal));

        const ULONGLONG now = GetTickCount64();
        ULONGLONG last      = _lastProgressTick.load(std::memory_order_acquire);
        if (! forcePost && now - last < kBatchRenameProgressThrottleMs)
        {
            return;
        }
        if (! _lastProgressTick.compare_exchange_strong(last, now, std::memory_order_acq_rel))
        {
            return;
        }

        _options.progressCallback(_options.progressContext, clampedCompleted, static_cast<uint64_t>(_overallTotal), forcePost);
    }

    const std::atomic_bool& _cancelRequested;
    BatchRenameExecutionOptions _options{};
    size_t _overallTotal = 0u;
    std::atomic<size_t> _progressBase{0u};
    std::atomic<ULONGLONG> _lastProgressTick{0ull};
    std::mutex _outcomesMutex;
    std::vector<BatchRenameItemOutcome> _itemOutcomes;
};

[[nodiscard]] HRESULT MakeBatchRenameTempLeaf(const std::filesystem::path& sourcePath, std::wstring& tempLeaf) noexcept
{
    // The temp leaf is "<name>.rsren-<32 hex>"; the suffix is a fixed 39 characters. Cap the name
    // portion so the synthesized leaf can never exceed the Windows 255-character limit -- otherwise a
    // cycle/swap of an already-long-named file could never be vacated (RenameItem would fail
    // ERROR_INVALID_NAME and strand the whole cycle). The GUID alone guarantees uniqueness; the name
    // prefix is only a readability aid, so truncating it is safe.
    constexpr size_t kMaxTempLeafLength    = 255u;
    const std::wstring sourceName = sourcePath.filename().native();
    return Common::Paths::BuildUniqueSiblingName(std::wstring_view(sourceName),
                                                  std::wstring_view(L".rsren-"),
                                                  std::wstring_view{},
                                                  kMaxTempLeafLength,
                                                  tempLeaf);
}

void FinalizeUndoEntryCurrentPaths(const FileSystemPathIdentity& pathIdentity,
                                   std::vector<BatchRenameUndoEntry>& undoEntries,
                                   std::span<const ExecutedDirectoryMove> directoryMoves)
{
    for (BatchRenameUndoEntry& entry : undoEntries)
    {
        entry.currentPath = ApplyExecutedDirectoryMoves(pathIdentity, entry.currentPath, directoryMoves);
    }
}

} // namespace

bool IsBatchRenameCancellationHRESULT(const HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED);
}

size_t PathDepthKey(const std::filesystem::path& path) noexcept
{
    size_t depth = 0u;
    for (const wchar_t ch : path.native())
    {
        if (ch == L'\\' || ch == L'/')
        {
            ++depth;
        }
    }
    return depth;
}

std::filesystem::path JoinFolderAndLeaf(const FileSystemPathIdentity& pathIdentity,
                                        const std::filesystem::path& folder,
                                        const std::wstring_view leaf) noexcept
{
    return std::filesystem::path(JoinFileSystemPath(pathIdentity, folder.native(), leaf));
}

std::filesystem::path ApplyExecutedDirectoryMoves(const FileSystemPathIdentity& pathIdentity,
                                                  std::filesystem::path path,
                                                  std::span<const ExecutedDirectoryMove> directoryMoves)
{
    for (const ExecutedDirectoryMove& move : directoryMoves)
    {
        // Rewrite strict descendants only. A path EQUAL to a move's source refers to
        // whatever occupies that name now (a sibling chain/swap rename re-occupying
        // the vacated name), not to the directory the move relocated; rewriting it
        // would corrupt undo entries and refresh targets for directory chains.
        if (IsStrictDescendantPath(pathIdentity, move.sourcePath.native(), path.native()))
        {
            path = std::filesystem::path(ReplaceFileSystemPathPrefix(pathIdentity, path.native(), move.sourcePath.native(), move.targetPath.native()));
        }
    }
    return path;
}

BatchRenameExecutionResult RunBatchRenameExecutionEngine(std::atomic_bool& cancelRequested,
                                                         IFileSystem& fileSystem,
                                                         const FileSystemPathIdentity pathIdentity,
                                                         std::vector<BatchRenameExecutionOp> ops,
                                                         const BatchRenameExecutionOptions options) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    BatchRenameExecutionCallback callback(cancelRequested, options, ops.size());
    BatchRenameExecutionResult result{};
    BatchRenameExecutionReport& report = result.report;
    report.totalRows                   = ops.size();

    size_t completedOps  = 0u;
    size_t failedOps     = 0u;
    HRESULT firstFailure = S_OK;
    HRESULT abortHr      = S_OK;
    bool aborted         = false;
    std::vector<ExecutedDirectoryMove> directoryMoves;
    directoryMoves.reserve(ops.size());

    const auto noteFailure = [&](const HRESULT status) noexcept
    {
        ++failedOps;
        if (SUCCEEDED(firstFailure))
        {
            firstFailure = status;
        }
    };

    const auto restoreTempBestEffort = [&](BatchRenameExecutionOp& op) noexcept
    {
        if (op.tempPath.empty() || op.completed || op.tempRestoreAttempted)
        {
            return;
        }

        op.tempRestoreAttempted = true;
        const HRESULT restoreHr = fileSystem.RenameItem(op.tempPath.c_str(), op.originalSource.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        if (FAILED(restoreHr))
        {
            Debug::Warning(L"BatchRename: failed to restore temporary rename '{}' back to '{}' (hr=0x{:08X}).",
                           op.tempPath.native(),
                           op.originalSource.native(),
                           static_cast<unsigned long>(restoreHr));
            return;
        }
        op.tempPath.clear();
    };

    const auto hasOpenTempHop = [&ops]() noexcept
    { return std::ranges::any_of(ops, [](const BatchRenameExecutionOp& op) noexcept { return ! op.tempPath.empty(); }); };

    const auto removeSuccessfulRecord = [&](const BatchRenameExecutionOp& op, const std::filesystem::path& targetPath)
    {
        for (size_t index = result.successfulSourcePaths.size(); index > 0u; --index)
        {
            const size_t entryIndex = index - 1u;
            if (EquivalentPath(pathIdentity, result.successfulSourcePaths[entryIndex].native(), op.originalSource.native()) &&
                entryIndex < result.successfulTargetPaths.size() &&
                EquivalentPath(pathIdentity, result.successfulTargetPaths[entryIndex].native(), targetPath.native()))
            {
                result.successfulSourcePaths.erase(result.successfulSourcePaths.begin() + static_cast<std::ptrdiff_t>(entryIndex));
                result.successfulTargetPaths.erase(result.successfulTargetPaths.begin() + static_cast<std::ptrdiff_t>(entryIndex));
                break;
            }
        }

        std::erase_if(report.undoEntries,
                      [&](const BatchRenameUndoEntry& entry) noexcept
        {
            return EquivalentPath(pathIdentity, entry.originalPath.native(), op.originalSource.native()) &&
                   EquivalentPath(pathIdentity, entry.currentPath.native(), targetPath.native());
        });
        std::erase_if(directoryMoves,
                      [&](const ExecutedDirectoryMove& move) noexcept
        {
            return EquivalentPath(pathIdentity, move.sourcePath.native(), op.originalSource.native()) &&
                   EquivalentPath(pathIdentity, move.targetPath.native(), targetPath.native());
        });
    };

    std::vector<size_t> completedWithOpenTempHop;
    completedWithOpenTempHop.reserve(ops.size());

    const auto rollbackCompletedOpenTempHopRows = [&]() noexcept
    {
        for (auto it = completedWithOpenTempHop.rbegin(); it != completedWithOpenTempHop.rend(); ++it)
        {
            BatchRenameExecutionOp& op = ops[*it];
            if (! op.completed)
            {
                continue;
            }

            const std::filesystem::path targetPath = JoinFolderAndLeaf(pathIdentity, op.currentSource.parent_path(), op.finalLeaf);
            const HRESULT rollbackHr = fileSystem.RenameItem(targetPath.c_str(), op.originalSource.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
            if (FAILED(rollbackHr))
            {
                Debug::Warning(L"BatchRename: failed to roll back completed rename '{}' back to '{}' while restoring a temp-hop cycle (hr=0x{:08X}).",
                               targetPath.native(),
                               op.originalSource.native(),
                               static_cast<unsigned long>(rollbackHr));
                continue;
            }

            removeSuccessfulRecord(op, targetPath);
            op.completed = false;
            op.failed    = true;
            if (completedOps != 0u)
            {
                --completedOps;
            }
            ++failedOps;
        }
        completedWithOpenTempHop.clear();
    };

    size_t groupBegin = 0u;
    while (groupBegin < ops.size() && ! aborted)
    {
        const size_t depth = ops[groupBegin].depth;
        size_t groupEnd    = groupBegin + 1u;
        while (groupEnd < ops.size() && ops[groupEnd].depth == depth)
        {
            ++groupEnd;
        }

        std::vector<size_t> remaining;
        remaining.reserve(groupEnd - groupBegin);
        for (size_t opIndex = groupBegin; opIndex < groupEnd; ++opIndex)
        {
            remaining.push_back(opIndex);
        }

        while (! remaining.empty() && ! aborted)
        {
            if (cancelRequested.load(std::memory_order_acquire))
            {
                aborted = true;
                abortHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                break;
            }

            std::vector<size_t> layer;
            layer.reserve(remaining.size());
            for (const size_t opIndex : remaining)
            {
                const std::filesystem::path targetPath = JoinFolderAndLeaf(pathIdentity, ops[opIndex].currentSource.parent_path(), ops[opIndex].finalLeaf);
                const bool blocked                     = std::ranges::any_of(remaining, [&](const size_t pendingIndex) noexcept {
                    return pendingIndex != opIndex && EquivalentPath(pathIdentity, ops[pendingIndex].currentSource.native(), targetPath.native());
                });
                if (! blocked)
                {
                    layer.push_back(opIndex);
                }
            }

            if (layer.empty())
            {
                // Every pending op targets another pending source: a rename
                // cycle (swap). Vacate one member through a unique temp name.
                const size_t cycleIndex         = remaining.front();
                BatchRenameExecutionOp& cycleOp = ops[cycleIndex];
                std::wstring tempLeaf;
                const HRESULT tempNameHr = MakeBatchRenameTempLeaf(cycleOp.currentSource, tempLeaf);
                if (FAILED(tempNameHr))
                {
                    cycleOp.failed = true;
                    noteFailure(tempNameHr);
                    std::erase(remaining, cycleIndex);
                    continue;
                }
                const FileSystemRenameBatch::RenameOp tempRename{
                    .sourcePath  = cycleOp.currentSource,
                    .newLeaf     = tempLeaf,
                    .depth       = cycleOp.depth,
                    .isDirectory = cycleOp.isDirectory,
                };
                callback.SetProgressBase(completedOps + failedOps);
                callback.ResetItemOutcomes();
                const HRESULT tempHr = FileSystemRenameBatch::Execute(
                    fileSystem, std::span<const FileSystemRenameBatch::RenameOp>(&tempRename, 1u), FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
                HRESULT tempStatus                                     = tempHr;
                const std::vector<BatchRenameItemOutcome> tempOutcomes = callback.TakeItemOutcomes();
                if (! tempOutcomes.empty())
                {
                    tempStatus = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    for (const BatchRenameItemOutcome& outcome : tempOutcomes)
                    {
                        if (outcome.itemIndex == 0u || EquivalentPath(pathIdentity, outcome.sourcePath, cycleOp.currentSource.native()))
                        {
                            tempStatus = outcome.status;
                            break;
                        }
                    }
                }

                if (FAILED(tempStatus))
                {
                    // The temp hop failed: surface it as this row's failure.
                    cycleOp.failed = true;
                    noteFailure(tempStatus);
                    std::erase(remaining, cycleIndex);
                    if (IsBatchRenameCancellationHRESULT(tempStatus))
                    {
                        aborted = true;
                        abortHr = tempStatus;
                    }
                    continue;
                }

                cycleOp.tempPath      = JoinFolderAndLeaf(pathIdentity, cycleOp.currentSource.parent_path(), tempLeaf);
                cycleOp.currentSource = cycleOp.tempPath;
                continue;
            }

            std::vector<FileSystemRenameBatch::RenameOp> batch;
            batch.reserve(layer.size());
            for (const size_t opIndex : layer)
            {
                batch.push_back(FileSystemRenameBatch::RenameOp{
                    .sourcePath  = ops[opIndex].currentSource,
                    .newLeaf     = ops[opIndex].finalLeaf,
                    .depth       = ops[opIndex].depth,
                    .isDirectory = ops[opIndex].isDirectory,
                });
            }

            callback.SetProgressBase(completedOps + failedOps);
            callback.ResetItemOutcomes();
            const HRESULT batchHr = FileSystemRenameBatch::Execute(fileSystem, batch, FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
            const std::vector<BatchRenameItemOutcome> outcomes = callback.TakeItemOutcomes();

            std::vector<std::optional<HRESULT>> statusByItemIndex(layer.size());
            for (const BatchRenameItemOutcome& outcome : outcomes)
            {
                if (outcome.itemIndex < statusByItemIndex.size())
                {
                    statusByItemIndex[outcome.itemIndex] = outcome.status;
                }
            }

            for (size_t layerOffset = 0u; layerOffset < layer.size(); ++layerOffset)
            {
                const size_t opIndex       = layer[layerOffset];
                BatchRenameExecutionOp& op = ops[opIndex];
                // Providers that report per-item completion drive per-row
                // bookkeeping; otherwise fall back to the batch result.
                HRESULT status = outcomes.empty() || IsBatchRenameCancellationHRESULT(batchHr) ? batchHr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                if (! outcomes.empty())
                {
                    if (statusByItemIndex[layerOffset].has_value())
                    {
                        status = statusByItemIndex[layerOffset].value();
                    }
                    else
                    {
                        for (const BatchRenameItemOutcome& outcome : outcomes)
                        {
                            if (outcome.itemIndex >= statusByItemIndex.size() && EquivalentPath(pathIdentity, outcome.sourcePath, op.currentSource.native()))
                            {
                                status = outcome.status;
                                break;
                            }
                        }
                    }
                }

                if (SUCCEEDED(status))
                {
                    op.completed = true;
                    ++completedOps;
                    const std::filesystem::path targetPath = JoinFolderAndLeaf(pathIdentity, op.currentSource.parent_path(), op.finalLeaf);
                    result.successfulSourcePaths.push_back(op.originalSource);
                    result.successfulTargetPaths.push_back(targetPath);
                    // Undo entries record the NET original -> final transition;
                    // temp hops are an implementation detail.
                    report.undoEntries.push_back(BatchRenameUndoEntry{
                        .currentPath  = targetPath,
                        .restoreName  = op.originalSource.filename().native(),
                        .originalPath = op.originalSource,
                    });
                    if (op.isDirectory)
                    {
                        directoryMoves.push_back(ExecutedDirectoryMove{.sourcePath = op.originalSource, .targetPath = targetPath});
                    }
                    op.tempPath.clear();
                    if (hasOpenTempHop())
                    {
                        completedWithOpenTempHop.push_back(opIndex);
                    }
                }
                else
                {
                    op.failed = true;
                    noteFailure(status);
                    restoreTempBestEffort(op);
                }
            }

            for (const size_t opIndex : layer)
            {
                std::erase(remaining, opIndex);
            }

            callback.NotifyOverallProgress(completedOps + failedOps);
            if (! hasOpenTempHop())
            {
                completedWithOpenTempHop.clear();
            }

            if (FAILED(batchHr))
            {
                if (hasOpenTempHop())
                {
                    rollbackCompletedOpenTempHopRows();
                }
                aborted = true;
                abortHr = batchHr;
            }
        }

        groupBegin = groupEnd;
    }

    // Ops that never ran (after a failed or canceled batch) count as failed,
    // and any active temp hop is rolled back best-effort.
    size_t neverRan = 0u;
    for (BatchRenameExecutionOp& op : ops)
    {
        if (op.completed || op.failed)
        {
            restoreTempBestEffort(op);
            continue;
        }
        ++neverRan;
        restoreTempBestEffort(op);
    }

    report.completedRows = completedOps;
    report.failedRows    = failedOps + neverRan;
    report.firstFailure  = (SUCCEEDED(firstFailure) && FAILED(abortHr)) ? abortHr : firstFailure;
    report.canceled      = IsBatchRenameCancellationHRESULT(report.firstFailure) || IsBatchRenameCancellationHRESULT(abortHr);
    FinalizeUndoEntryCurrentPaths(pathIdentity, report.undoEntries, directoryMoves);

    HRESULT terminalHr       = S_OK;
    std::wstring_view detail = L"success";
    if (report.failedRows != 0u || FAILED(report.firstFailure))
    {
        terminalHr = FAILED(report.firstFailure) ? report.firstFailure : E_FAIL;
        detail     = report.canceled ? L"canceled" : L"rename_failed";
    }

    result.hr                     = terminalHr;
    result.detail                 = detail;
    result.executedDirectoryMoves = std::move(directoryMoves);

    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"batchrename.execute.us",
                          detail,
                          Debug::Perf::ElapsedUs(startedAt),
                          static_cast<uint64_t>(ops.size()),
                          static_cast<uint64_t>(completedOps),
                          terminalHr);
        Debug::Perf::EmitValue(L"batchrename.execute.rows", static_cast<uint64_t>(ops.size()));
        Debug::Perf::EmitValue(L"batchrename.execute.completed", static_cast<uint64_t>(completedOps));
        Debug::Perf::EmitValue(L"batchrename.execute.failed", static_cast<uint64_t>(report.failedRows));
    }

    return result;
}
