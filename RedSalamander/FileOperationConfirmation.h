#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"

#include <filesystem>
#include <string>
#include <string_view>
#include <vector>

struct NonRevertableFileOperationPromptCounts
{
    unsigned long long fileCount    = 0;
    unsigned long long folderCount  = 0;
    unsigned long long unknownCount = 0;
    std::filesystem::path sampleFile;
    bool hasSampleFile = false;
};

[[nodiscard]] bool ConfirmNonRevertableFileOperation(HWND owner,
                                                     IFileSystem* fileSystem,
                                                     FileSystemOperation operation,
                                                     const std::vector<std::filesystem::path>& sourcePaths,
                                                     const std::filesystem::path& destinationFolder) noexcept;

[[nodiscard]] bool ConfirmNonRevertableFileOperation(HWND owner,
                                                     FileSystemOperation operation,
                                                     const std::vector<std::filesystem::path>& sourcePaths,
                                                     const std::filesystem::path& destinationFolder,
                                                     const NonRevertableFileOperationPromptCounts& counts,
                                                     HostFileOperationPromptOptions* options = nullptr,
                                                     std::wstring_view confirmationMessage   = {}) noexcept;
