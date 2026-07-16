#pragma once

#include <cstddef>
#include <filesystem>
#include <span>
#include <string>

#include "PlugInterfaces/FileSystem.h"

namespace FileSystemRenameBatch
{
struct RenameOp final
{
    std::filesystem::path sourcePath;
    std::wstring newLeaf;
    size_t depth     = 0u;
    bool isDirectory = false;
};

[[nodiscard]] HRESULT Execute(IFileSystem& fileSystem,
                              std::span<const RenameOp> ops,
                              FileSystemFlags flags,
                              const FileSystemOptions* options = nullptr,
                              IFileSystemCallback* callback    = nullptr,
                              void* cookie                     = nullptr) noexcept;
} // namespace FileSystemRenameBatch
