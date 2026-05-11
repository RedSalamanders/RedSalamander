#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>

#include <filesystem>
#include <string>
#include <vector>

namespace RedConfigure::Workspace
{
struct ResourceOwner
{
    std::wstring name;
    std::filesystem::path projectPath;
    std::filesystem::path embeddedResourcePath;
    std::vector<std::filesystem::path> satelliteResourcePaths;
};

struct ThemeFile
{
    std::filesystem::path path;
};

struct WorkspaceScanResult
{
    std::filesystem::path root;
    std::vector<ResourceOwner> resourceOwners;
    std::vector<ThemeFile> themeFiles;
    std::vector<std::wstring> errors;
};

HRESULT DiscoverWorkspace(const std::filesystem::path& root, WorkspaceScanResult& outResult) noexcept;
} // namespace RedConfigure::Workspace
