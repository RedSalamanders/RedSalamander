#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>

#include "SettingsStore.h"
#include "Workspace/WorkspaceDiscovery.h"

#include <filesystem>
#include <span>
#include <string>
#include <vector>

namespace RedConfigure::Themes
{
struct ThemeCatalogEntry
{
    std::filesystem::path path;
    Common::Settings::ThemeDefinition definition;
};

struct ThemeCatalog
{
    std::vector<ThemeCatalogEntry> themes;
    std::vector<std::wstring> errors;
};

HRESULT LoadThemeCatalog(std::span<const Workspace::ThemeFile> files, ThemeCatalog& outCatalog) noexcept;
} // namespace RedConfigure::Themes
