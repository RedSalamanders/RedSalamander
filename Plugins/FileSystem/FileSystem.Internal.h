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
} // namespace FileSystemInternal
