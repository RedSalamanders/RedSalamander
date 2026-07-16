#pragma once

#include "FileSystem.h"
#include "Helpers.h"

namespace FileSystemInternal
{
struct PathInfo
{
    std::wstring display;
    std::wstring extended;
};

std::wstring MakeAbsolutePath(const std::wstring& path);
std::wstring ToExtendedPath(const std::wstring& path);

bool TryGetUncServerRoot(std::wstring_view path, std::wstring& serverName) noexcept;

[[nodiscard]] bool IsDotOrDotDot(std::wstring_view name) noexcept;

std::wstring AppendPath(const std::wstring& base, std::wstring_view leaf);
std::wstring AppendPath(const std::wstring& base, const wchar_t* leaf);

std::wstring_view TrimTrailingSeparators(std::wstring_view path) noexcept;
std::wstring_view GetPathLeaf(std::wstring_view path) noexcept;
std::wstring GetPathDirectory(std::wstring_view path);
[[nodiscard]] bool ContainsPathSeparator(std::wstring_view text) noexcept;

PathInfo MakePathInfo(const std::wstring& path);
PathInfo MakePathInfo(const wchar_t* path);

struct StagedPromotionOptions
{
    bool allowReplaceReadOnly        = false;
    bool preserveReplacementReadOnly = false;
    bool replacementIsReparsePoint   = false;
    bool stripTemporaryAttributes    = false;
    bool ignoreReplaceMergeErrors    = true;
};

[[nodiscard]] HRESULT PromoteStagedTempIntoFinalPath(const std::wstring& tempPath,
                                                     const std::wstring& finalPath,
                                                     const StagedPromotionOptions& options) noexcept;

// Module anchor for AcquireModuleReferenceFromAddress — keeps the DLL loaded
// while background worker threads or threadpool callbacks are active.
extern const int kFileSystemModuleAnchor;

// Stops and joins the shared background copy/move worker threads.
// Intended to be invoked at a host "quiet point" when the last FileSystem instance is being destroyed.
void ShutdownSharedFileOpsJobScheduler() noexcept;

#if defined(_DEBUG)
void RunDebugPathNormalizationSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugReparseCopyErrorMappingSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugDirectorySizeErrorPolicySelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugSharedFileOpsSchedulerShutdownSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugSearchServiceFallbackCandidateSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
#endif
} // namespace FileSystemInternal
