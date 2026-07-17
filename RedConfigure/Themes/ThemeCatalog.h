#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>

#include "SettingsStore.h"
#include "Workspace/WorkspaceDiscovery.h"

#include <cstdint>
#include <filesystem>
#include <span>
#include <string>
#include <vector>

namespace RedConfigure::Themes
{
enum class ThemeCatalogOrigin : uint8_t
{
    BuiltIn,
    File,
    User,
};

struct ThemeCatalogEntry
{
    std::filesystem::path path;
    Common::Settings::ThemeDefinition definition;
    ThemeCatalogOrigin origin = ThemeCatalogOrigin::File;
};

struct ThemeCatalog
{
    std::vector<ThemeCatalogEntry> themes;
    std::vector<std::wstring> errors;
};

HRESULT LoadThemeCatalog(std::span<const Workspace::ThemeFile> files, ThemeCatalog& outCatalog) noexcept;
} // namespace RedConfigure::Themes
