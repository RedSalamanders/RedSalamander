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
    std::filesystem::path providerDestinationPath;
    std::wstring providerParentPath;
    std::wstring providerParentKey;
    std::wstring providerCollisionKey;
    size_t depth     = 0u;
    bool isDirectory = false;
};

// Callers must populate each operation from a successful provider child-name
// contract and revalidate that contract immediately before this mutation boundary.
[[nodiscard]] HRESULT Execute(IFileSystem& fileSystem,
                              std::span<const RenameOp> ops,
                              FileSystemFlags flags,
                              const FileSystemOptions* options = nullptr,
                              IFileSystemCallback* callback    = nullptr,
                              void* cookie                     = nullptr) noexcept;
} // namespace FileSystemRenameBatch
