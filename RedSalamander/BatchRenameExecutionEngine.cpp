#include "Framework.h"

#include "BatchRenameExecutionEngine.h"

#include "Helpers.h"

#include <algorithm>
#include <chrono>
#include <format>
#include <limits>
#include <numeric>
#include <optional>
#include <unordered_map>
#include <utility>

namespace
{
constexpr ULONGLONG kBatchRenameProgressThrottleMs = 100ull;

class BatchRenameProgressSink final
{
public:
    BatchRenameProgressSink(BatchRenameExecutionOptions options, const size_t overallTotal) noexcept
        : _options(options),
          _overallTotal(overallTotal)
    {
    }

    BatchRenameProgressSink(const BatchRenameProgressSink&)            = delete;
    BatchRenameProgressSink& operator=(const BatchRenameProgressSink&) = delete;
    BatchRenameProgressSink(BatchRenameProgressSink&&)                 = delete;
    BatchRenameProgressSink& operator=(BatchRenameProgressSink&&)      = delete;

    void NotifyOverallProgress(const size_t processedOps) noexcept
    {
        PostProgress(static_cast<uint64_t>(processedOps), processedOps >= _overallTotal);
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

    BatchRenameExecutionOptions _options{};
    size_t _overallTotal = 0u;
    std::atomic<ULONGLONG> _lastProgressTick{0ull};
};

void FinalizeUndoEntryCurrentPaths(const FileSystemPathIdentity& pathIdentity,
                                   std::vector<BatchRenameUndoEntry>& undoEntries,
                                   std::span<const ExecutedDirectoryMove> directoryMoves)
{
    for (BatchRenameUndoEntry& entry : undoEntries)
    {
        entry.currentPath = ApplyExecutedDirectoryMoves(pathIdentity, entry.currentPath, directoryMoves);
    }
}

[[nodiscard]] bool ValidateExecutionSchedule(const BatchRenameExecutionSchedule& schedule, const size_t operationCount) noexcept
{
    if (! schedule.cycleOperationIndices.empty())
    {
        return false;
    }
    std::vector<bool> seen(operationCount, false);
    size_t covered = 0u;
    for (const BatchRenameExecutionLayer& layer : schedule.layers)
    {
        if (layer.operationIndices.empty())
        {
            return false;
        }
        for (const size_t index : layer.operationIndices)
        {
            if (index >= seen.size() || seen[index])
            {
                return false;
            }
            seen[index] = true;
            ++covered;
        }
    }
    return covered == operationCount;
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
        // whatever occupies that name now, not to the directory the move relocated.
        if (IsStrictDescendantPath(pathIdentity, move.sourcePath.native(), path.native()))
        {
            path = std::filesystem::path(
                ReplaceFileSystemPathPrefix(pathIdentity, path.native(), move.sourcePath.native(), move.targetPath.native()));
        }
    }
    return path;
}

HRESULT BuildBatchRenameExecutionSchedule(const FileSystemPathIdentity& /*pathIdentity*/,
                                          const std::span<const BatchRenameExecutionOp> operations,
                                          BatchRenameExecutionSchedule& out) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    out                  = {};
    uint64_t retainedIndexBytes = 0u;
    const auto finish = [&](const HRESULT hr) noexcept
    {
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"batchrename.schedule.build.us",
                              FAILED(hr) ? L"rejected" : L"accepted",
                              Debug::Perf::ElapsedUs(startedAt),
                              static_cast<uint64_t>(operations.size()),
                              static_cast<uint64_t>(out.layers.size()),
                              hr);
            Debug::Perf::EmitValue(L"batchrename.schedule.layers", static_cast<uint64_t>(out.layers.size()), hr);
            Debug::Perf::EmitValue(L"batchrename.schedule.cycle_members", static_cast<uint64_t>(out.cycleOperationIndices.size()), hr);
            Debug::Perf::EmitValue(L"batchrename.schedule.retained_index_bytes", retainedIndexBytes, hr);
        }
        return hr;
    };

    if (operations.empty())
    {
        return finish(S_OK);
    }
    for (const BatchRenameExecutionOp& operation : operations)
    {
        if (operation.originalSource.empty() || operation.finalLeaf.empty() || operation.providerFinalPath.empty() ||
            operation.providerParentKey.empty() || operation.providerSourceCollisionKey.empty() || operation.providerFinalCollisionKey.empty())
        {
            return finish(E_INVALIDARG);
        }
    }

    std::vector<size_t> orderedIndices(operations.size());
    std::iota(orderedIndices.begin(), orderedIndices.end(), 0u);
    std::ranges::sort(orderedIndices, [&](const size_t left, const size_t right) noexcept
    {
        if (operations[left].depth != operations[right].depth)
        {
            return operations[left].depth > operations[right].depth;
        }
        return left < right;
    });

    std::vector<std::optional<size_t>> dependency(operations.size());
    size_t groupBegin = 0u;
    while (groupBegin < orderedIndices.size())
    {
        const size_t depth = operations[orderedIndices[groupBegin]].depth;
        size_t groupEnd    = groupBegin + 1u;
        while (groupEnd < orderedIndices.size() && operations[orderedIndices[groupEnd]].depth == depth)
        {
            ++groupEnd;
        }

        const auto makeLocationKey = [](const std::wstring_view parentKey, const std::wstring_view childKey)
        {
            std::wstring key;
            key.reserve(parentKey.size() + 1u + childKey.size());
            key.append(parentKey);
            key.push_back(L'\0');
            key.append(childKey);
            return key;
        };

        std::unordered_map<std::wstring, std::vector<size_t>> sourcesByKey;
        sourcesByKey.reserve(groupEnd - groupBegin);
        for (size_t position = groupBegin; position < groupEnd; ++position)
        {
            const size_t index = orderedIndices[position];
            std::wstring key = makeLocationKey(operations[index].providerParentKey, operations[index].providerSourceCollisionKey);
            retainedIndexBytes += static_cast<uint64_t>(key.capacity() + 1u) * sizeof(wchar_t) + sizeof(size_t);
            sourcesByKey[std::move(key)].push_back(index);
        }

        for (size_t position = groupBegin; position < groupEnd; ++position)
        {
            const size_t index = orderedIndices[position];
            if (operations[index].providerFinalCollisionKey == operations[index].providerSourceCollisionKey)
            {
                continue;
            }

            const std::wstring destinationKey =
                makeLocationKey(operations[index].providerParentKey, operations[index].providerFinalCollisionKey);
            if (const auto candidates = sourcesByKey.find(destinationKey); candidates != sourcesByKey.end())
            {
                for (const size_t candidate : candidates->second)
                {
                    if (candidate != index)
                    {
                        dependency[index] = candidate;
                        break;
                    }
                }
            }
        }

        std::vector<bool> remaining(operations.size(), false);
        size_t remainingCount = groupEnd - groupBegin;
        for (size_t position = groupBegin; position < groupEnd; ++position)
        {
            remaining[orderedIndices[position]] = true;
        }

        while (remainingCount != 0u)
        {
            BatchRenameExecutionLayer layer{};
            layer.operationIndices.reserve(remainingCount);
            for (size_t position = groupBegin; position < groupEnd; ++position)
            {
                const size_t index = orderedIndices[position];
                if (! remaining[index])
                {
                    continue;
                }
                if (! dependency[index].has_value() || ! remaining[dependency[index].value()])
                {
                    layer.operationIndices.push_back(index);
                }
            }
            if (layer.operationIndices.empty())
            {
                std::vector<uint8_t> visit(operations.size(), 0u);
                std::vector<bool> cycleMember(operations.size(), false);
                for (size_t position = groupBegin; position < groupEnd; ++position)
                {
                    const size_t start = orderedIndices[position];
                    if (! remaining[start] || visit[start] != 0u)
                    {
                        continue;
                    }
                    std::vector<size_t> stack;
                    size_t current = start;
                    while (remaining[current] && visit[current] == 0u)
                    {
                        visit[current] = 1u;
                        stack.push_back(current);
                        if (! dependency[current].has_value() || ! remaining[dependency[current].value()])
                        {
                            break;
                        }
                        current = dependency[current].value();
                    }
                    if (remaining[current] && visit[current] == 1u)
                    {
                        const auto cycleBegin = std::ranges::find(stack, current);
                        for (auto cycle = cycleBegin; cycle != stack.end(); ++cycle)
                        {
                            cycleMember[*cycle] = true;
                        }
                    }
                    for (const size_t visited : stack)
                    {
                        visit[visited] = 2u;
                    }
                }
                for (size_t index = 0u; index < cycleMember.size(); ++index)
                {
                    if (cycleMember[index])
                    {
                        out.cycleOperationIndices.push_back(index);
                    }
                }
                out.layers.clear();
                return finish(HRESULT_FROM_WIN32(ERROR_CIRCULAR_DEPENDENCY));
            }

            for (const size_t index : layer.operationIndices)
            {
                remaining[index] = false;
                --remainingCount;
            }
            out.layers.push_back(std::move(layer));
        }
        groupBegin = groupEnd;
    }

    return finish(ValidateExecutionSchedule(out, operations.size()) ? S_OK : E_UNEXPECTED);
}

BatchRenameExecutionResult RunBatchRenameExecutionEngine(std::atomic_bool& cancelRequested,
                                                         const FileSystemPathIdentity pathIdentity,
                                                         std::vector<BatchRenameExecutionOp> ops,
                                                         const BatchRenameExecutionOptions options) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    BatchRenameProgressSink progress(options, ops.size());
    BatchRenameExecutionResult result{};
    BatchRenameExecutionReport& report = result.report;
    report.totalRows                   = ops.size();
    uint64_t mutationCount             = 0u;

    const auto finish = [&](const HRESULT terminalHr, const std::wstring_view detail) -> BatchRenameExecutionResult
    {
        result.hr     = terminalHr;
        result.detail = detail;
        result.itemResults.clear();
        result.itemResults.reserve(ops.size());
        for (size_t index = 0u; index < ops.size(); ++index)
        {
            const BatchRenameExecutionOp& op = ops[index];
            result.itemResults.push_back(BatchRenameExecutionItemResult{
                .operationIndex = index,
                .status         = op.status,
                .completed      = op.completed,
                .skipped        = op.skipped,
            });
        }
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"batchrename.execute.us",
                              detail,
                              Debug::Perf::ElapsedUs(startedAt),
                              static_cast<uint64_t>(ops.size()),
                              static_cast<uint64_t>(report.completedRows),
                              terminalHr);
            Debug::Perf::EmitValue(L"batchrename.execute.rows", static_cast<uint64_t>(ops.size()));
            Debug::Perf::EmitValue(L"batchrename.execute.completed", static_cast<uint64_t>(report.completedRows));
            Debug::Perf::EmitValue(L"batchrename.execute.failed", static_cast<uint64_t>(report.failedRows));
            Debug::Perf::EmitValue(L"batchrename.execute.mutations", mutationCount, terminalHr);
            Debug::Perf::EmitValue(L"batchrename.execute.journal_writes", 0u, terminalHr);
        }
        return std::move(result);
    };

    if (options.mutationCallback == nullptr)
    {
        report.failedRows   = ops.size();
        report.firstFailure = E_INVALIDARG;
        for (BatchRenameExecutionOp& op : ops)
        {
            op.failed = true;
            op.status = E_INVALIDARG;
        }
        return finish(E_INVALIDARG, L"missing_mutation_owner");
    }

    BatchRenameExecutionSchedule builtSchedule{};
    const BatchRenameExecutionSchedule* schedule = options.schedule;
    if (schedule == nullptr)
    {
        const HRESULT scheduleHr = BuildBatchRenameExecutionSchedule(pathIdentity, ops, builtSchedule);
        if (FAILED(scheduleHr))
        {
            report.failedRows   = ops.size();
            report.firstFailure = scheduleHr;
            for (BatchRenameExecutionOp& op : ops)
            {
                op.failed = true;
                op.status = scheduleHr;
            }
            return finish(scheduleHr,
                          scheduleHr == HRESULT_FROM_WIN32(ERROR_CIRCULAR_DEPENDENCY) ? L"dependency_cycle" : L"invalid_schedule");
        }
        schedule = &builtSchedule;
    }
    if (! ValidateExecutionSchedule(*schedule, ops.size()))
    {
        report.failedRows   = ops.size();
        report.firstFailure = E_INVALIDARG;
        for (BatchRenameExecutionOp& op : ops)
        {
            op.failed = true;
            op.status = E_INVALIDARG;
        }
        return finish(E_INVALIDARG, L"invalid_schedule");
    }

    size_t completedOps  = 0u;
    size_t skippedOps    = 0u;
    size_t failedOps     = 0u;
    HRESULT firstFailure = S_OK;
    HRESULT abortHr      = S_OK;
    bool aborted         = false;
    std::vector<ExecutedDirectoryMove> directoryMoves;
    directoryMoves.reserve(ops.size());

    for (const BatchRenameExecutionLayer& layer : schedule->layers)
    {
        if (aborted)
        {
            break;
        }
        if (cancelRequested.load(std::memory_order_acquire))
        {
            aborted = true;
            abortHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            break;
        }

        HRESULT layerHr = S_OK;
        for (const size_t opIndex : layer.operationIndices)
        {
            BatchRenameExecutionOp& op = ops[opIndex];
            if (cancelRequested.load(std::memory_order_acquire))
            {
                layerHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                break;
            }

            const std::filesystem::path& targetPath = op.providerFinalPath;
            ++mutationCount;
            const HRESULT status = options.mutationCallback(options.mutationContext, opIndex, op.originalSource, targetPath);
            if (status == S_FALSE)
            {
                op.skipped = true;
                op.status  = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                ++skippedOps;
                if (layerHr == S_OK)
                {
                    layerHr = S_FALSE;
                }
                continue;
            }
            if (SUCCEEDED(status))
            {
                op.completed = true;
                op.status    = S_OK;
                ++completedOps;
                result.successfulSourcePaths.push_back(op.originalSource);
                result.successfulTargetPaths.push_back(targetPath);
                report.undoEntries.push_back(BatchRenameUndoEntry{
                    .currentPath  = targetPath,
                    .restoreName  = op.originalSource.filename().native(),
                    .originalPath = op.originalSource,
                });
                if (op.isDirectory)
                {
                    directoryMoves.push_back(ExecutedDirectoryMove{.sourcePath = op.originalSource, .targetPath = targetPath});
                }
                continue;
            }

            op.failed = true;
            op.status = status;
            ++failedOps;
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = status;
            }
            if (SUCCEEDED(layerHr))
            {
                layerHr = status;
            }
            if (IsBatchRenameCancellationHRESULT(status))
            {
                break;
            }
        }

        progress.NotifyOverallProgress(completedOps + skippedOps + failedOps);
        if (layerHr != S_OK)
        {
            aborted = true;
            abortHr = layerHr;
        }
    }

    size_t neverRan = 0u;
    for (BatchRenameExecutionOp& op : ops)
    {
        if (op.completed || op.skipped || op.failed)
        {
            continue;
        }
        ++neverRan;
        op.failed = true;
        op.status = FAILED(abortHr) ? abortHr : HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED);
    }

    report.completedRows = completedOps;
    report.skippedRows   += skippedOps;
    report.failedRows     = failedOps + neverRan;
    report.firstFailure   = (SUCCEEDED(firstFailure) && FAILED(abortHr)) ? abortHr : firstFailure;
    report.canceled       = IsBatchRenameCancellationHRESULT(report.firstFailure) || IsBatchRenameCancellationHRESULT(abortHr);
    FinalizeUndoEntryCurrentPaths(pathIdentity, report.undoEntries, directoryMoves);
    result.executedDirectoryMoves = std::move(directoryMoves);

    if (report.failedRows != 0u || FAILED(report.firstFailure))
    {
        const HRESULT terminalHr = FAILED(report.firstFailure) ? report.firstFailure : E_FAIL;
        return finish(terminalHr, report.canceled ? L"canceled" : L"rename_failed");
    }
    if (report.skippedRows != 0u)
    {
        return finish(S_FALSE, L"partial");
    }
    progress.NotifyOverallProgress(ops.size());
    return finish(S_OK, L"success");
}
