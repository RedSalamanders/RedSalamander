#include "ThemeCatalog.h"

#include "ThemeDefinitionIo.h"

#include <algorithm>
#include <fstream>
#include <iterator>

namespace
{
[[nodiscard]] bool ReadFileText(const std::filesystem::path& path, std::string& outText) noexcept
{
    outText.clear();
    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return false;
    }
    outText.assign(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
    return ! input.bad();
}
} // namespace

namespace RedConfigure::Themes
{
HRESULT LoadThemeCatalog(std::span<const Workspace::ThemeFile> files, ThemeCatalog& outCatalog) noexcept
{
    outCatalog = {};
    for (const Workspace::ThemeFile& file : files)
    {
        std::string json;
        if (! ReadFileText(file.path, json))
        {
            outCatalog.errors.push_back(file.path.wstring());
            continue;
        }

        Common::Settings::ThemeDefinition definition;
        Common::Settings::ThemeDefinitionIoError error = Common::Settings::ThemeDefinitionIoError::None;
        if (FAILED(Common::Settings::ParseThemeDefinitionJson5(json, definition, &error, nullptr)))
        {
            outCatalog.errors.push_back(file.path.wstring());
            continue;
        }

        const ThemeCatalogOrigin origin = definition.id.starts_with(L"builtin/") ? ThemeCatalogOrigin::BuiltIn : ThemeCatalogOrigin::File;
        outCatalog.themes.push_back(ThemeCatalogEntry{.path = file.path, .definition = std::move(definition), .origin = origin});
    }

    std::sort(outCatalog.themes.begin(),
              outCatalog.themes.end(),
              [](const auto& lhs, const auto& rhs) noexcept
    {
        if (lhs.definition.name == rhs.definition.name)
        {
            return lhs.definition.id < rhs.definition.id;
        }
        return lhs.definition.name < rhs.definition.name;
    });

    return S_OK;
}
} // namespace RedConfigure::Themes
