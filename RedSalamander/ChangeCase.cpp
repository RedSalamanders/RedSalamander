#include "ChangeCase.h"

#include "framework.h"

#include <algorithm>
#include <filesystem>
#include <iterator>
#include <optional>
#include <span>
#include <string>
#include <unordered_set>
#include <vector>

#include <wil/com.h>

#ifdef ENABLE_TESTS
#include "FileSystemRenameBatch.h"
#endif
#include "FileSystemRouteContract.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "PlugInterfaces/FileSystem.h"

namespace
{
[[nodiscard]] std::wstring ToLower(std::wstring_view text) noexcept
{
    std::wstring out(text);
    if (! out.empty())
    {
        ::CharLowerBuffW(out.data(), static_cast<DWORD>(out.size()));
    }
    return out;
}

[[nodiscard]] std::wstring ToUpper(std::wstring_view text) noexcept
{
    std::wstring out(text);
    if (! out.empty())
    {
        ::CharUpperBuffW(out.data(), static_cast<DWORD>(out.size()));
    }
    return out;
}

[[nodiscard]] std::wstring ToMixed(std::wstring_view text) noexcept
{
    std::wstring out(text);
    if (out.empty())
    {
        return out;
    }

    ::CharLowerBuffW(out.data(), static_cast<DWORD>(out.size()));

    bool newWord = true;
    for (wchar_t& ch : out)
    {
        // Classify and upcase through the Windows character tables so non-ASCII initials
        // (e.g. "école") capitalize correctly; the C-locale towupper/iswalnum are ASCII-only.
        if (::IsCharAlphaNumericW(ch) == FALSE)
        {
            newWord = true;
            continue;
        }

        if (newWord)
        {
            ::CharUpperBuffW(&ch, 1u);
            newWord = false;
        }
    }

    return out;
}

[[nodiscard]] size_t ChangeCasePathDepthKey(const std::filesystem::path& p) noexcept
{
    size_t depth             = 0;
    const std::wstring& text = p.native();
    for (wchar_t ch : text)
    {
        if (ch == L'\\' || ch == L'/')
        {
            ++depth;
        }
    }
    return depth;
}

[[nodiscard]] HRESULT ClassifyDirectoryReadResult(const HRESULT readHr,
                                                  const bool hasInformation,
                                                  const bool /*provenDirectory*/) noexcept
{
    if (FAILED(readHr))
    {
        return readHr;
    }
    return hasInformation ? S_OK : E_UNEXPECTED;
}

[[nodiscard]] bool IsDotOrDotDot(std::wstring_view name) noexcept
{
    return name == L"." || name == L"..";
}

} // namespace

namespace ChangeCase
{
#ifdef ENABLE_TESTS
HRESULT DebugClassifyDirectoryReadResult(const HRESULT readHr,
                                         const bool hasInformation,
                                         const bool provenDirectory) noexcept
{
    return ClassifyDirectoryReadResult(readHr, hasInformation, provenDirectory);
}
#endif

std::wstring TransformLeafName(std::wstring_view leafName, const Options& options) noexcept
{
    const std::filesystem::path leafPath{std::wstring(leafName)};
    const std::wstring stem = leafPath.stem().wstring();
    const std::wstring ext  = leafPath.extension().wstring(); // includes leading dot

    switch (options.target)
    {
        case ChangeTarget::WholeFilename:
        {
            switch (options.style)
            {
                case CaseStyle::Lower: return ToLower(leafName);
                case CaseStyle::Upper: return ToUpper(leafName);
                case CaseStyle::Mixed: return ToMixed(leafName);
                case CaseStyle::PartiallyMixed:
                {
                    std::wstring result = ToMixed(stem);
                    result.append(ToLower(ext));
                    return result;
                }
            }
            break;
        }
        case ChangeTarget::OnlyName:
        {
            std::wstring result;
            switch (options.style)
            {
                case CaseStyle::Lower: result = ToLower(stem); break;
                case CaseStyle::Upper: result = ToUpper(stem); break;
                case CaseStyle::Mixed: result = ToMixed(stem); break;
                case CaseStyle::PartiallyMixed: result = ToMixed(stem); break;
            }
            result.append(ext);
            return result;
        }
        case ChangeTarget::OnlyExtension:
        {
            if (ext.empty())
            {
                return std::wstring(leafName);
            }

            std::wstring newExt;
            switch (options.style)
            {
                case CaseStyle::Lower: newExt = ToLower(ext); break;
                case CaseStyle::Upper: newExt = ToUpper(ext); break;
                case CaseStyle::Mixed: newExt = ToMixed(ext); break;
                case CaseStyle::PartiallyMixed: newExt = ToLower(ext); break;
            }

            std::wstring result = stem;
            result.append(std::move(newExt));
            return result;
        }
    }

    return std::wstring(leafName);
}

[[nodiscard]] HRESULT BuildOrApplyToPaths(IFileSystem& fileSystem,
                                          const std::wstring_view pluginId,
                                          const std::vector<std::filesystem::path>& inputPaths,
                                          const Options& options,
                                          std::vector<BatchRenameExecutionOp>* const operationsOut,
                                          std::stop_token stopToken,
                                          ProgressCallback progress,
                                          void* progressCookie,
#ifdef ENABLE_TESTS
                                          const MutationGuardCallbacks* const mutationGuard) noexcept
#else
                                          const void*) noexcept
#endif
{
    if (inputPaths.empty())
    {
        return S_OK;
    }
    if (pluginId.empty())
    {
        return E_INVALIDARG;
    }

    const auto firstPath = std::ranges::find_if(inputPaths, [](const std::filesystem::path& path) noexcept { return ! path.empty(); });
    if (firstPath == inputPaths.end())
    {
        return E_INVALIDARG;
    }
    wil::com_ptr<IFileSystemRouteCapabilities> route;
    const HRESULT routeInterfaceHr = fileSystem.QueryInterface(__uuidof(IFileSystemRouteCapabilities), route.put_void());
    if (FAILED(routeInterfaceHr) || ! route)
    {
        return FAILED(routeInterfaceHr) ? routeInterfaceHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    const FileSystemRouteContract::QueryResult routeResult =
        FileSystemRouteContract::Query(route.get(), firstPath->native(), FILESYSTEM_RENAME, pluginId);
    if (routeResult.state != FileSystemRouteContract::QueryState::Available || ! routeResult.snapshot.renameOperation ||
        ! routeResult.snapshot.pathIdentity.has_value())
    {
        return FAILED(routeResult.status) ? routeResult.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    const FileSystemPathIdentity pathIdentity = routeResult.snapshot.pathIdentity.value();
    const auto childNameFailure = [](const FileSystemRouteContract::ChildNameContractResult& contract) noexcept
    {
        if (contract.state == FileSystemRouteContract::QueryState::Available && contract.nameStatus == FILESYSTEM_CHILD_NAME_INVALID &&
            FAILED(contract.failureStatus))
        {
            return contract.failureStatus;
        }
        return FAILED(contract.status) ? contract.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    };

    ProgressUpdate progressUpdate{};

    std::vector<std::filesystem::path> paths;
    paths.reserve(inputPaths.size());

    std::unordered_set<std::wstring> seen;
    seen.reserve(inputPaths.size() * 2u);

    const auto addPath = [&](const std::filesystem::path& p) noexcept -> HRESULT
    {
        if (p.empty())
        {
            return S_FALSE;
        }

        const std::optional<std::wstring> key = TryMakePathKey(pathIdentity, p.native());
        if (! key.has_value())
        {
            return E_INVALIDARG;
        }

        if (seen.insert(key.value()).second)
        {
            paths.push_back(p);
        }
        return S_OK;
    };

    for (const auto& p : inputPaths)
    {
        const HRESULT addHr = addPath(p);
        if (FAILED(addHr))
        {
            return addHr;
        }
    }

    if (options.includeSubdirs)
    {
        struct PendingDirectory final
        {
            std::filesystem::path path;
            bool provenDirectory = false;
        };

        std::vector<PendingDirectory> pending;
        pending.reserve(inputPaths.size());
        for (const auto& root : inputPaths)
        {
            constexpr FileSystemBindFlags selectedRootBindFlags = static_cast<FileSystemBindFlags>(
                FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_RENAME);
            FileOperations::ObjectBindingResult selectedRoot = FileOperations::BindObjectAuthority(
                &fileSystem, root.native(), routeResult.snapshot.pathProfileId, selectedRootBindFlags);
            if (selectedRoot.state != FileOperations::ObjectBindingState::Bound)
            {
                return FAILED(selectedRoot.status) ? selectedRoot.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }
            if (selectedRoot.authority.kind == FILESYSTEM_BOUND_DIRECTORY)
            {
                pending.push_back(PendingDirectory{.path = root, .provenDirectory = true});
            }
        }

        while (! pending.empty())
        {
            if (stopToken.stop_requested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            PendingDirectory candidate = std::move(pending.back());
            pending.pop_back();
            std::filesystem::path& directory = candidate.path;

            progressUpdate.phase       = ProgressUpdate::Phase::Enumerating;
            progressUpdate.currentPath = directory;
            if (progress)
            {
                progress(progressUpdate, progressCookie);
            }

            wil::com_ptr<IFilesInformation> info;
            const HRESULT readHr = fileSystem.ReadDirectoryInfo(directory.c_str(), info.addressof());
            const HRESULT readDisposition =
                ClassifyDirectoryReadResult(readHr, static_cast<bool>(info), candidate.provenDirectory);
            if (FAILED(readDisposition))
            {
                return readDisposition;
            }

            FileInfo* buffer = nullptr;
            HRESULT hr       = info->GetBuffer(&buffer);
            if (FAILED(hr))
            {
                return hr;
            }

            if (! buffer)
            {
                ++progressUpdate.scannedFolders;
                continue;
            }

            unsigned long bufferSize = 0;
            hr                       = info->GetBufferSize(&bufferSize);
            if (FAILED(hr))
            {
                return hr;
            }

            unsigned char* bytes = reinterpret_cast<unsigned char*>(buffer);
            unsigned long offset = 0;

            for (FileInfo* entry = buffer; entry;)
            {
                if (stopToken.stop_requested())
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                if ((entry->FileNameSize % sizeof(wchar_t)) != 0u)
                {
                    return E_INVALIDARG;
                }

                const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
                std::wstring_view name(entry->FileName, nameChars);
                if (! name.empty() && ! IsDotOrDotDot(name))
                {
                    // Existing children only need the provider join; validation applies to the
                    // proposed target names below, so a legacy name never blocks discovery.
                    const FileSystemRouteContract::StringResult childJoined =
                        FileSystemRouteContract::QueryJoinedPath(route.get(), directory.native(), name, FILESYSTEM_RENAME);
                    if (childJoined.state != FileSystemRouteContract::QueryState::Available || childJoined.value.empty())
                    {
                        return FAILED(childJoined.status) ? childJoined.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                    }
                    const std::filesystem::path child(childJoined.value);
                    const std::optional<std::wstring> key = TryMakePathKey(pathIdentity, child.native());
                    if (! key.has_value())
                    {
                        return E_INVALIDARG;
                    }
                    if (seen.insert(key.value()).second)
                    {
                        paths.push_back(child);

                        const DWORD attrs         = entry->FileAttributes;
                        const bool isDir          = (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0;
                        const bool isReparsePoint = (attrs & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                        if (isDir && ! isReparsePoint)
                        {
                            pending.push_back(PendingDirectory{.path = child, .provenDirectory = true});
                        }
                    }
                }

                ++progressUpdate.scannedEntries;

                if (entry->NextEntryOffset == 0)
                {
                    break;
                }

                if (entry->NextEntryOffset > bufferSize - offset)
                {
                    return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
                }

                offset += entry->NextEntryOffset;
                if (offset >= bufferSize)
                {
                    return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
                }

                entry = reinterpret_cast<FileInfo*>(bytes + offset);
            }

            ++progressUpdate.scannedFolders;
        }
    }

    if (paths.size() > 1u)
    {
        std::ranges::sort(paths,
                          [](const std::filesystem::path& a, const std::filesystem::path& b) noexcept
        {
            const size_t depthA = ChangeCasePathDepthKey(a);
            const size_t depthB = ChangeCasePathDepthKey(b);
            if (depthA != depthB)
            {
                return depthA > depthB;
            }
            return a.native().size() > b.native().size();
        });
    }

    std::vector<BatchRenameExecutionOp> renames;
    renames.reserve(paths.size());
    std::unordered_set<std::wstring> targetLocations;
    targetLocations.reserve(paths.size());

    for (const auto& path : paths)
    {
        std::wstring parentPath;
        std::wstring oldLeaf;
        if (! TryGetFileSystemParentPath(pathIdentity, path.native(), parentPath) ||
            ! TryGetFileSystemLeafName(pathIdentity, path.native(), oldLeaf))
        {
            return E_INVALIDARG;
        }

        const std::wstring newLeaf = TransformLeafName(oldLeaf, options);
        const std::optional<std::wstring> parentKey = TryMakePathKey(pathIdentity, parentPath);
        if (! parentKey.has_value())
        {
            return E_INVALIDARG;
        }
        const FileSystemRouteContract::StringResult sourceCollisionKey =
            FileSystemRouteContract::QueryChildNameCollisionKey(route.get(), parentPath, oldLeaf, FILESYSTEM_RENAME);
        if (sourceCollisionKey.state != FileSystemRouteContract::QueryState::Available ||
            FAILED(sourceCollisionKey.status) || sourceCollisionKey.value.empty())
        {
            return FAILED(sourceCollisionKey.status) ? sourceCollisionKey.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        std::wstring targetCollisionKey;
        std::optional<FileSystemRouteContract::ChildNameContractResult> nameContract;
        if (newLeaf == oldLeaf)
        {
            targetCollisionKey = sourceCollisionKey.value;
        }
        else
        {
            nameContract = FileSystemRouteContract::QueryChildNameContract(
                route.get(), parentPath, newLeaf, FILESYSTEM_RENAME, pluginId);
            if (nameContract->state != FileSystemRouteContract::QueryState::Available ||
                nameContract->nameStatus != FILESYSTEM_CHILD_NAME_VALID || nameContract->joinedPath.empty() ||
                nameContract->collisionKey.empty())
            {
                return childNameFailure(nameContract.value());
            }
            targetCollisionKey = nameContract->collisionKey;
        }

        std::wstring targetLocation;
        targetLocation.reserve(parentKey->size() + 1u + targetCollisionKey.size());
        targetLocation.append(parentKey.value());
        targetLocation.push_back(L'\0');
        targetLocation.append(targetCollisionKey);
        if (! targetLocations.insert(std::move(targetLocation)).second)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        if (! nameContract.has_value())
        {
            continue;
        }

        BatchRenameExecutionOp op{};
        op.originalSource = path;
        op.finalLeaf    = std::move(newLeaf);
        op.providerFinalPath = std::filesystem::path(nameContract->joinedPath);
        op.providerParentKey = parentKey.value();
        op.providerSourceCollisionKey = sourceCollisionKey.value;
        op.providerFinalCollisionKey = nameContract->collisionKey;
        op.depth      = ChangeCasePathDepthKey(path);
        renames.push_back(std::move(op));
    }

    progressUpdate.plannedRenames = static_cast<uint64_t>(renames.size());
    if (operationsOut != nullptr)
    {
        *operationsOut = std::move(renames);
        if (progress)
        {
            progress(progressUpdate, progressCookie);
        }
        return S_OK;
    }

#ifndef ENABLE_TESTS
    // Production callers always request immutable operations and return above. Direct provider
    // mutation remains compiled only into self-test builds while E5 retires the old authority.
    return E_UNEXPECTED;
#else
    if (mutationGuard != nullptr && mutationGuard->prepare != nullptr && ! renames.empty())
    {
        std::vector<std::filesystem::path> mutationPaths;
        mutationPaths.reserve(renames.size());
        std::ranges::transform(renames, std::back_inserter(mutationPaths), [](const BatchRenameExecutionOp& op)
        {
            return op.originalSource;
        });
        const HRESULT guardHr = mutationGuard->prepare(mutationPaths, mutationGuard->cookie);
        if (guardHr != S_OK)
        {
            return guardHr == S_FALSE ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : guardHr;
        }
    }

    progressUpdate.phase = ProgressUpdate::Phase::Renaming;
    progressUpdate.currentPath.clear();
    progressUpdate.plannedRenames   = static_cast<uint64_t>(renames.size());
    progressUpdate.completedRenames = 0;
    if (progress)
    {
        progress(progressUpdate, progressCookie);
    }

    uint64_t completed          = 0;
    size_t index                = 0;
    const FileSystemFlags flags = FILESYSTEM_FLAG_NONE;
    while (index < renames.size())
    {
        if (stopToken.stop_requested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        const size_t depth = renames[index].depth;
        size_t end         = index + 1u;
        while (end < renames.size() && renames[end].depth == depth)
        {
            ++end;
        }

        constexpr size_t kBatchSize = 64u;
        for (size_t batchStart = index; batchStart < end; batchStart += kBatchSize)
        {
            if (stopToken.stop_requested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            const size_t batchEnd = std::min(end, batchStart + kBatchSize);

            if (mutationGuard != nullptr && mutationGuard->revalidate != nullptr)
            {
                std::vector<std::filesystem::path> mutationPaths;
                mutationPaths.reserve(batchEnd - batchStart);
                std::ranges::transform(
                    std::span<const BatchRenameExecutionOp>(renames.data() + batchStart, batchEnd - batchStart),
                    std::back_inserter(mutationPaths),
                    [](const BatchRenameExecutionOp& op)
                    {
                        return op.originalSource;
                    });
                const HRESULT guardHr = mutationGuard->revalidate(mutationPaths, mutationGuard->cookie);
                if (guardHr != S_OK)
                {
                    return guardHr == S_FALSE ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : guardHr;
                }
            }

            for (size_t operationIndex = batchStart; operationIndex < batchEnd; ++operationIndex)
            {
                const BatchRenameExecutionOp& operation = renames[operationIndex];
                std::wstring providerParentPath;
                if (! TryGetFileSystemParentPath(pathIdentity, operation.originalSource.native(), providerParentPath))
                {
                    return E_INVALIDARG;
                }
                const FileSystemRouteContract::ChildNameContractResult current = FileSystemRouteContract::QueryChildNameContract(
                    route.get(), providerParentPath, operation.finalLeaf, FILESYSTEM_RENAME, pluginId);
                if (current.state != FileSystemRouteContract::QueryState::Available ||
                    current.nameStatus != FILESYSTEM_CHILD_NAME_VALID || current.joinedPath != operation.providerFinalPath.native() ||
                    current.collisionKey != operation.providerFinalCollisionKey)
                {
                    if (current.state == FileSystemRouteContract::QueryState::Available &&
                        current.nameStatus == FILESYSTEM_CHILD_NAME_INVALID && FAILED(current.failureStatus))
                    {
                        return current.failureStatus;
                    }
                    return FAILED(current.status) ? current.status : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
                }
            }

            progressUpdate.currentPath = renames[batchStart].originalSource;
            if (progress)
            {
                progress(progressUpdate, progressCookie);
            }

            std::vector<FileSystemRenameBatch::RenameOp> legacyBatch;
            legacyBatch.reserve(batchEnd - batchStart);
            for (size_t operationIndex = batchStart; operationIndex < batchEnd; ++operationIndex)
            {
                const BatchRenameExecutionOp& operation = renames[operationIndex];
                std::wstring providerParentPath;
                if (! TryGetFileSystemParentPath(pathIdentity, operation.originalSource.native(), providerParentPath))
                {
                    return E_INVALIDARG;
                }
                legacyBatch.emplace_back(FileSystemRenameBatch::RenameOp{
                    .sourcePath = operation.originalSource,
                    .newLeaf = operation.finalLeaf,
                    .providerDestinationPath = operation.providerFinalPath,
                    .providerParentPath = std::move(providerParentPath),
                    .providerParentKey = operation.providerParentKey,
                    .providerCollisionKey = operation.providerFinalCollisionKey,
                    .depth = operation.depth,
                });
            }
            const HRESULT hr = FileSystemRenameBatch::Execute(fileSystem, legacyBatch, flags);
            if (FAILED(hr))
            {
                return hr;
            }

            completed += static_cast<uint64_t>(batchEnd - batchStart);
            progressUpdate.completedRenames = completed;
            progressUpdate.currentPath      = renames[batchEnd - 1u].originalSource;
            if (progress)
            {
                progress(progressUpdate, progressCookie);
            }
        }

        index = end;
    }

    return S_OK;
#endif
}

HRESULT BuildRenameOperations(IFileSystem& fileSystem,
                              const std::wstring_view pluginId,
                              const std::vector<std::filesystem::path>& inputPaths,
                              const Options& options,
                              std::vector<BatchRenameExecutionOp>& operationsOut,
                              std::stop_token stopToken,
                              ProgressCallback progress,
                              void* progressCookie) noexcept
{
    operationsOut.clear();
    return BuildOrApplyToPaths(
        fileSystem, pluginId, inputPaths, options, &operationsOut, stopToken, progress, progressCookie, nullptr);
}

#ifdef ENABLE_TESTS
HRESULT DebugApplyToPathsForTests(IFileSystem& fileSystem,
                                  const std::wstring_view pluginId,
                                  const std::vector<std::filesystem::path>& inputPaths,
                                  const Options& options,
                                  std::stop_token stopToken,
                                  ProgressCallback progress,
                                  void* progressCookie,
                                  const MutationGuardCallbacks* const mutationGuard) noexcept
{
    return BuildOrApplyToPaths(
        fileSystem, pluginId, inputPaths, options, nullptr, stopToken, progress, progressCookie, mutationGuard);
}
#endif
} // namespace ChangeCase
