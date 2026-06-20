#include "FileSystemRenameBatch.h"

#include <cstdint>
#include <limits>
#include <string_view>

namespace
{
[[nodiscard]] bool ContainsPathSeparator(std::wstring_view text) noexcept
{
    return text.find(L'\\') != std::wstring_view::npos || text.find(L'/') != std::wstring_view::npos;
}

[[nodiscard]] bool IsUnsupportedBulkRename(HRESULT hr) noexcept
{
    return hr == E_NOTIMPL || hr == HRESULT_FROM_WIN32(ERROR_CALL_NOT_IMPLEMENTED) || hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

[[nodiscard]] HRESULT ExecuteRenameItemFallback(IFileSystem& fileSystem,
                                                std::span<const FileSystemRenameBatch::RenameOp> ops,
                                                FileSystemFlags flags,
                                                const FileSystemOptions* options,
                                                IFileSystemCallback* callback,
                                                void* cookie) noexcept
{
    for (size_t index = 0; index < ops.size(); ++index)
    {
        if (callback)
        {
            BOOL cancel = FALSE;
            if (SUCCEEDED(callback->FileSystemShouldCancel(&cancel, cookie)) && cancel != FALSE)
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
        }

        const FileSystemRenameBatch::RenameOp& op   = ops[index];
        const std::filesystem::path destinationPath = op.sourcePath.parent_path() / op.newLeaf;
        const HRESULT hr = fileSystem.RenameItem(op.sourcePath.c_str(), destinationPath.c_str(), flags, options, callback, cookie);
        if (callback)
        {
            // Report each attempted rename (including the failing one) so hosts that track
            // per-item outcomes see exactly which items reached the filesystem; otherwise a
            // partial fallback run would be indistinguishable from a total failure.
            static_cast<void>(callback->FileSystemItemCompleted(
                FILESYSTEM_RENAME, static_cast<unsigned long>(index), op.sourcePath.c_str(), destinationPath.c_str(), hr, nullptr, cookie));
        }
        if (FAILED(hr))
        {
            return hr;
        }
    }

    return S_OK;
}
} // namespace

namespace FileSystemRenameBatch
{
HRESULT Execute(IFileSystem& fileSystem,
                std::span<const RenameOp> ops,
                FileSystemFlags flags,
                const FileSystemOptions* options,
                IFileSystemCallback* callback,
                void* cookie) noexcept
{
    if (ops.empty())
    {
        return S_OK;
    }

    uint64_t totalBytes64 = static_cast<uint64_t>(ops.size()) * static_cast<uint64_t>(sizeof(FileSystemRenamePair));
    for (const RenameOp& op : ops)
    {
        const std::wstring& sourceText = op.sourcePath.native();
        const size_t sourceLen         = sourceText.size();
        const size_t nameLen           = op.newLeaf.size();

        totalBytes64 += static_cast<uint64_t>((sourceLen + 1u) * sizeof(wchar_t));
        totalBytes64 += static_cast<uint64_t>((nameLen + 1u) * sizeof(wchar_t));
        if (totalBytes64 > static_cast<uint64_t>((std::numeric_limits<unsigned long>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
    }

    FileSystemArenaOwner arenaOwner;
    const HRESULT initHr = arenaOwner.Initialize(static_cast<unsigned long>(totalBytes64));
    if (FAILED(initHr))
    {
        return initHr;
    }

    FileSystemArena* arena = arenaOwner.Get();
    auto* pairs            = static_cast<FileSystemRenamePair*>(AllocateFromFileSystemArena(
        arena, static_cast<unsigned long>(ops.size() * sizeof(FileSystemRenamePair)), static_cast<unsigned long>(alignof(FileSystemRenamePair))));
    if (! pairs)
    {
        return E_OUTOFMEMORY;
    }

    for (size_t i = 0; i < ops.size(); ++i)
    {
        const RenameOp& op           = ops[i];
        const std::wstring& source   = op.sourcePath.native();
        const std::wstring_view name = op.newLeaf;

        if (source.empty() || name.empty())
        {
            return E_INVALIDARG;
        }

        if (ContainsPathSeparator(name))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const size_t sourceLen = source.size();
        const size_t nameLen   = name.size();

        if (sourceLen > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()) - 1u ||
            nameLen > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()) - 1u)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        auto* sourceBuf = static_cast<wchar_t*>(
            AllocateFromFileSystemArena(arena, static_cast<unsigned long>((sourceLen + 1u) * sizeof(wchar_t)), static_cast<unsigned long>(alignof(wchar_t))));
        if (! sourceBuf)
        {
            return E_OUTOFMEMORY;
        }

        auto* nameBuf = static_cast<wchar_t*>(
            AllocateFromFileSystemArena(arena, static_cast<unsigned long>((nameLen + 1u) * sizeof(wchar_t)), static_cast<unsigned long>(alignof(wchar_t))));
        if (! nameBuf)
        {
            return E_OUTOFMEMORY;
        }

        ::CopyMemory(sourceBuf, source.data(), sourceLen * sizeof(wchar_t));
        sourceBuf[sourceLen] = L'\0';

        ::CopyMemory(nameBuf, name.data(), nameLen * sizeof(wchar_t));
        nameBuf[nameLen] = L'\0';

        pairs[i].sizeBytes  = sizeof(FileSystemRenamePair);
        pairs[i].sourcePath = sourceBuf;
        pairs[i].newName    = nameBuf;
    }

    if (ops.size() > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const HRESULT bulkHr = fileSystem.RenameItems(pairs, static_cast<unsigned long>(ops.size()), flags, options, callback, cookie);
    if (! IsUnsupportedBulkRename(bulkHr))
    {
        return bulkHr;
    }

    return ExecuteRenameItemFallback(fileSystem, ops, flags, options, callback, cookie);
}
} // namespace FileSystemRenameBatch
