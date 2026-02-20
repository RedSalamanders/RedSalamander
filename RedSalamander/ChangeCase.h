#pragma once

#include "framework.h"

#include <cstdint>
#include <filesystem>
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
    CaseStyle style      = CaseStyle::Lower;
    ChangeTarget target  = ChangeTarget::WholeFilename;
    bool includeSubdirs  = false;
};

[[nodiscard]] std::wstring TransformLeafName(std::wstring_view leafName, const Options& options) noexcept;

struct ProgressUpdate final
{
    enum class Phase : uint8_t
    {
        Enumerating,
        Renaming,
    };

    Phase phase = Phase::Enumerating;
    std::filesystem::path currentPath;

    uint64_t scannedFolders   = 0;
    uint64_t scannedEntries   = 0;

    uint64_t plannedRenames   = 0;
    uint64_t completedRenames = 0;
};

using ProgressCallback = void (*)(const ProgressUpdate& update, void* cookie) noexcept;

// Applies the requested case transformation to the given paths.
// Notes:
// - includeSubdirs uses IFileSystem::ReadDirectoryInfo (non-recursive traversal).
// - Renames are batched via IFileSystem::RenameItems.
// - stopToken allows cooperative cancellation.
[[nodiscard]] HRESULT ApplyToPaths(IFileSystem& fileSystem,
                                  const std::vector<std::filesystem::path>& inputPaths,
                                  const Options& options,
                                  std::stop_token stopToken = {},
                                  ProgressCallback progress = nullptr,
                                  void* progressCookie      = nullptr) noexcept;
} // namespace ChangeCase
