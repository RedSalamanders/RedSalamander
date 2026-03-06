#pragma once

#include <filesystem>

namespace AppDataPaths
{
// Resolves %LOCALAPPDATA% via SHGetKnownFolderPath first; falls back to
// GetEnvironmentVariableW in case the shell API is unavailable.
[[nodiscard]] std::filesystem::path GetLocalAppDataPath() noexcept;
} // namespace AppDataPaths

