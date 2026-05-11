#include "RedConfigureApp.h"
#include "Localization/PlaceholderValidation.h"
#include "Localization/RcParser.h"
#include "Localization/RcWriter.h"
#include "RedConfigureSession.h"
#include "SettingsStore.h"
#include "Themes/ThemeCatalog.h"
#include "Themes/ThemePreviewModel.h"
#include "ThemeDefinitionIo.h"
#include "Workspace/WorkspaceDiscovery.h"

#include <array>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <string>
#include <string_view>
#include <system_error>

namespace
{
[[nodiscard]] bool Require(bool condition, std::wstring_view message)
{
    if (! condition)
    {
        std::wcerr << message << L'\n';
        return false;
    }

    return true;
}

[[nodiscard]] bool WriteTestTextFile(const std::filesystem::path& path, std::string_view text)
{
    std::error_code ec;
    std::filesystem::create_directories(path.parent_path(), ec);
    if (ec)
    {
        return false;
    }

    std::ofstream output(path, std::ios::binary);
    if (! output)
    {
        return false;
    }

    output.write(text.data(), static_cast<std::streamsize>(text.size()));
    return output.good();
}

[[nodiscard]] bool WriteTestUtf16LeTextFile(const std::filesystem::path& path, std::wstring_view text, bool bom)
{
    std::error_code ec;
    std::filesystem::create_directories(path.parent_path(), ec);
    if (ec)
    {
        return false;
    }

    std::ofstream output(path, std::ios::binary);
    if (! output)
    {
        return false;
    }

    if (bom)
    {
        constexpr std::array<uint8_t, 2u> marker{{0xFFu, 0xFEu}};
        output.write(reinterpret_cast<const char*>(marker.data()), static_cast<std::streamsize>(marker.size()));
    }
    if (! text.empty())
    {
        output.write(reinterpret_cast<const char*>(text.data()), static_cast<std::streamsize>(text.size() * sizeof(wchar_t)));
    }
    return output.good();
}

[[nodiscard]] bool TestPageDefinitions()
{
    const auto pages = RedConfigure::GetPageDefinitions();
    bool ok          = true;
    ok               = Require(pages.size() == 4u, L"RedConfigure should expose the four task-mode pages.") && ok;

    constexpr std::array expectedIds = {
        std::wstring_view(L"start"),
        std::wstring_view(L"localization"),
        std::wstring_view(L"themes"),
        std::wstring_view(L"reviewExport"),
    };

    if (pages.size() == expectedIds.size())
    {
        for (size_t index = 0; index < expectedIds.size(); ++index)
        {
            ok = Require(pages[index].id == expectedIds[index], L"RedConfigure page id order changed unexpectedly.") && ok;
            ok = Require(pages[index].titleResourceId != 0u, L"RedConfigure page title resource id must be non-zero.") && ok;
            ok = Require(pages[index].descriptionResourceId != 0u, L"RedConfigure page description resource id must be non-zero.") && ok;
        }
    }

    return ok;
}

[[nodiscard]] bool TestResolveWorkspaceRootForLaunchPath()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(ec) / L"RedConfigureWorkspaceRootResolveTest";
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for workspace-root tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    ec.clear();

    bool ok = true;
    ok      = Require(WriteTestTextFile(tempRoot / L"RedSalamander.sln", "\n"), L"Failed to write solution marker fixture.") && ok;
    std::filesystem::create_directories(tempRoot / L".build" / L"x64" / L"Debug", ec);
    ok = Require(! ec, L"Failed to create nested build folder fixture.") && ok;

    const std::filesystem::path resolved = RedConfigure::ResolveWorkspaceRootForLaunchPath(tempRoot / L".build" / L"x64" / L"Debug");
    ok = Require(resolved == tempRoot, L"Launch paths inside .build should resolve back to the repository root.") && ok;

    const std::filesystem::path unknownRoot = tempRoot.parent_path() / L"RedConfigureStandaloneRootResolveTest";
    const std::filesystem::path unknownResolved = RedConfigure::ResolveWorkspaceRootForLaunchPath(unknownRoot);
    ok = Require(unknownResolved == unknownRoot, L"Paths without repository markers should be preserved.") && ok;

    std::filesystem::remove_all(tempRoot, ec);
    std::filesystem::remove_all(unknownRoot, ec);
    return ok;
}

[[nodiscard]] bool TestWorkspaceDiscoveryFindsResourcesAndThemes()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(ec) / L"RedConfigureWorkspaceDiscoveryTest";
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for workspace discovery tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    ec.clear();

    bool ok = true;
    ok      = Require(WriteTestTextFile(tempRoot / L"RedSalamander" / L"RedSalamander.vcxproj",
                                   R"xml(<?xml version="1.0" encoding="utf-8"?>
<Project>
  <ItemGroup>
    <ResourceCompile Include="RedSalamander.rc" />
  </ItemGroup>
</Project>)xml"),
                 L"Failed to write app project fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"RedSalamander" / L"RedSalamander.rc", "STRINGTABLE\nBEGIN\nEND\n"),
                 L"Failed to write app rc fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"RedSalamander" / L"Lang" / L"fr-FR" / L"RedSalamander-fr-FR.rc", "STRINGTABLE\nBEGIN\nEND\n"),
                 L"Failed to write satellite rc fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"Plugins" / L"ViewerText" / L"ViewerText.vcxproj",
                                   R"xml(<?xml version="1.0" encoding="utf-8"?>
<Project>
  <ItemGroup>
    <ResourceCompile Include="ViewerText.rc" />
  </ItemGroup>
</Project>)xml"),
                 L"Failed to write plugin project fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"Plugins" / L"ViewerText" / L"ViewerText.rc", "STRINGTABLE\nBEGIN\nEND\n"),
                 L"Failed to write plugin rc fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"Specs" / L"Themes" / L"Forest.theme.json5", "{ id: 'user/forest' }\n"),
                 L"Failed to write theme fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L".build" / L"x64" / L"Debug" / L"Ignored.vcxproj",
                                   R"xml(<Project><ItemGroup><ResourceCompile Include="Ignored.rc" /></ItemGroup></Project>)xml"),
                 L"Failed to write ignored build project fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L".build" / L"x64" / L"Debug" / L"Themes" / L"Ignored.theme.json5", "{}\n"),
                 L"Failed to write ignored build theme fixture.") &&
         ok;

    RedConfigure::Workspace::WorkspaceScanResult result;
    const HRESULT hr = RedConfigure::Workspace::DiscoverWorkspace(tempRoot, result);
    ok               = Require(SUCCEEDED(hr), L"Workspace discovery should succeed for the fixture repo.") && ok;
    ok               = Require(result.errors.empty(), L"Workspace discovery should not report fixture errors.") && ok;
    ok               = Require(result.resourceOwners.size() == 2u, L"Workspace discovery should find the app and plugin resource owners.") && ok;
    ok               = Require(result.themeFiles.size() == 1u, L"Workspace discovery should find only the non-build theme file.") && ok;

    const auto ownerByName = [&](std::wstring_view name) noexcept -> const RedConfigure::Workspace::ResourceOwner*
    {
        for (const auto& owner : result.resourceOwners)
        {
            if (owner.name == name)
            {
                return &owner;
            }
        }
        return nullptr;
    };

    const auto* appOwner = ownerByName(L"RedSalamander");
    ok                  = Require(appOwner != nullptr, L"Workspace discovery should expose the RedSalamander owner.") && ok;
    if (appOwner)
    {
        ok = Require(appOwner->embeddedResourcePath.filename() == L"RedSalamander.rc", L"App owner should expose its embedded resource file.") && ok;
        ok = Require(appOwner->satelliteResourcePaths.size() == 1u, L"App owner should expose one satellite resource file.") && ok;
    }

    const auto* pluginOwner = ownerByName(L"ViewerText");
    ok                     = Require(pluginOwner != nullptr, L"Workspace discovery should expose the ViewerText owner.") && ok;
    if (pluginOwner)
    {
        ok = Require(pluginOwner->embeddedResourcePath.filename() == L"ViewerText.rc", L"Plugin owner should expose its embedded resource file.") && ok;
    }

    if (! result.themeFiles.empty())
    {
        ok = Require(result.themeFiles.front().path.filename() == L"Forest.theme.json5", L"Workspace discovery should expose the expected theme file.") && ok;
    }

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestWorkspaceDiscoveryReportsScanErrors()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(ec) / L"RedConfigureWorkspaceDiscoveryErrorTest";
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for workspace discovery error tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    ec.clear();

    bool ok = true;
    ok      = Require(WriteTestTextFile(tempRoot / L"Broken" / L"Broken.vcxproj", "<Project><ItemGroup><ResourceCompile Include=\"Broken.rc\"></Project>"),
                 L"Failed to write broken project fixture.") &&
         ok;

    RedConfigure::Workspace::WorkspaceScanResult result;
    const HRESULT hr = RedConfigure::Workspace::DiscoverWorkspace(tempRoot, result);
    ok               = Require(SUCCEEDED(hr), L"Workspace discovery should return a result even when one project cannot be parsed.") && ok;
    ok               = Require(! result.errors.empty(), L"Workspace discovery should surface project parse errors.") && ok;

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestThemeDefinitionJson5Parsing()
{
    constexpr std::string_view json = R"json5(
{
  // JSON5 comments and trailing commas are accepted.
  "id": "user/forest-test",
  "name": "Forest Test",
  "baseThemeId": "builtin/dark",
  "colors": {
    "app.accent": "#2ECC71",
    "folderView.itemBackgroundSelected": "#802ECC71",
  },
}
)json5";

    Common::Settings::ThemeDefinition theme;
    Common::Settings::ThemeDefinitionIoError error = Common::Settings::ThemeDefinitionIoError::None;
    bool ok                                        = true;
    const HRESULT hr                               = Common::Settings::ParseThemeDefinitionJson5(json, theme, &error, nullptr);
    ok                                             = Require(SUCCEEDED(hr), L"Expected JSON5 theme parsing to succeed.") && ok;
    ok                                             = Require(error == Common::Settings::ThemeDefinitionIoError::None, L"Expected no theme parse error.") && ok;
    ok                                             = Require(theme.id == L"user/forest-test", L"Parsed theme id did not match.") && ok;
    ok                                             = Require(theme.name == L"Forest Test", L"Parsed theme name did not match.") && ok;
    ok                                             = Require(theme.baseThemeId == L"builtin/dark", L"Parsed base theme did not match.") && ok;
    ok                                             = Require(theme.colors.size() == 2u, L"Parsed theme color count did not match.") && ok;
    ok = Require(theme.colors[L"app.accent"] == 0xFF2ECC71u, L"Parsed #RRGGBB color did not get opaque alpha.") && ok;
    ok = Require(theme.colors[L"folderView.itemBackgroundSelected"] == 0x802ECC71u, L"Parsed #AARRGGBB color did not match.") && ok;

    return ok;
}

[[nodiscard]] bool TestThemeDefinitionRejectsInvalidInput()
{
    Common::Settings::ThemeDefinition theme;
    Common::Settings::ThemeDefinitionIoError error = Common::Settings::ThemeDefinitionIoError::None;
    bool ok                                        = true;

    const HRESULT emptyInput = Common::Settings::ParseThemeDefinitionJson5("", theme, &error, nullptr);
    ok = Require(emptyInput == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Empty theme input should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::EmptyInput, L"Empty theme input should report the empty-input error.") && ok;

    const HRESULT missingId = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"name":"Bad","baseThemeId":"builtin/dark","colors":{}})json", theme, &error, nullptr);
    ok = Require(missingId == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Missing id should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::MissingOrInvalidId, L"Missing id should report the id error.") && ok;

    const HRESULT missingName = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"id":"user/bad","baseThemeId":"builtin/dark","colors":{}})json", theme, &error, nullptr);
    ok = Require(missingName == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Missing name should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::MissingOrInvalidName, L"Missing name should report the name error.") && ok;

    const HRESULT missingBaseTheme = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"id":"user/bad","name":"Bad","colors":{}})json", theme, &error, nullptr);
    ok = Require(missingBaseTheme == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Missing base theme should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::MissingOrInvalidBaseThemeId, L"Missing base theme should report the base-theme error.") && ok;

    const HRESULT missingColors = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"id":"user/bad","name":"Bad","baseThemeId":"builtin/dark"})json", theme, &error, nullptr);
    ok = Require(missingColors == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Missing colors should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::ColorsMissingOrNotObject, L"Missing colors should report the colors error.") && ok;

    const HRESULT invalidId = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"id":"builtin/dark","name":"Bad","baseThemeId":"builtin/dark","colors":{}})json", theme, &error, nullptr);
    ok = Require(invalidId == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Invalid user theme id should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::InvalidId, L"Invalid id should report the id error.") && ok;

    const HRESULT invalidBaseTheme = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"id":"user/bad","name":"Bad","baseThemeId":"user/base","colors":{}})json", theme, &error, nullptr);
    ok = Require(invalidBaseTheme == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Invalid base theme should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::InvalidBaseThemeId, L"Invalid base theme should report the base-theme error.") && ok;

    const HRESULT invalidColorKey = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"id":"user/bad","name":"Bad","baseThemeId":"builtin/dark","colors":{"bad key":"#112233"}})json", theme, &error, nullptr);
    ok = Require(invalidColorKey == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Invalid color key should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::InvalidColorKey, L"Invalid color key should report the color-key error.") && ok;

    const HRESULT colorNotString = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"id":"user/bad","name":"Bad","baseThemeId":"builtin/dark","colors":{"app.accent":12}})json", theme, &error, nullptr);
    ok = Require(colorNotString == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Non-string color should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::ColorValueNotString, L"Non-string color should report the color-string error.") && ok;

    const HRESULT invalidColor = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"id":"user/bad","name":"Bad","baseThemeId":"builtin/dark","colors":{"app.accent":"red"}})json", theme, &error, nullptr);
    ok = Require(invalidColor == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Invalid color should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::InvalidColorValue, L"Invalid color should report the color value error.") && ok;

    return ok;
}

[[nodiscard]] bool TestThemeDefinitionJson5Export()
{
    Common::Settings::ThemeDefinition theme;
    theme.id          = L"user/export-test";
    theme.name        = L"Export Test";
    theme.baseThemeId = L"builtin/light";
    theme.colors.emplace(L"folderView.background", 0xFFFFFFFFu);
    theme.colors.emplace(L"app.accent", 0xFF0055AAu);

    std::string json;
    bool ok          = true;
    const HRESULT hr = Common::Settings::BuildThemeDefinitionJson5(theme, json);
    ok               = Require(SUCCEEDED(hr), L"Expected theme export to succeed.") && ok;
    ok               = Require(json.find("\"id\": \"user/export-test\"") != std::string::npos, L"Exported theme id is missing.") && ok;
    ok               = Require(json.find("\"app.accent\": \"#0055AA\"") < json.find("\"folderView.background\": \"#FFFFFF\""),
                 L"Exported colors should be sorted by key.") &&
         ok;

    Common::Settings::ThemeDefinition reparsed;
    Common::Settings::ThemeDefinitionIoError error = Common::Settings::ThemeDefinitionIoError::None;
    const HRESULT parseHr                          = Common::Settings::ParseThemeDefinitionJson5(json, reparsed, &error, nullptr);
    ok                                             = Require(SUCCEEDED(parseHr), L"Exported theme should parse again.") && ok;
    ok                                             = Require(reparsed.colors == theme.colors, L"Reparsed exported colors did not round trip.") && ok;

    return ok;
}

[[nodiscard]] bool TestSettingsStoreThemeDirectoryUsesSharedParser()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(ec) / L"RedConfigureSettingsThemeParserParityTest";
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for SettingsStore theme parser parity tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    bool ok = true;
    ok      = Require(WriteTestTextFile(tempRoot / L"Good.theme.json5",
                                   R"json5({
  id: "user/settings-good",
  name: "Settings Good",
  baseThemeId: "builtin/dark",
  colors: {
    "app.accent": "#336699",
  },
})json5"),
                 L"Failed to write valid SettingsStore theme fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"InvalidId.theme.json5",
                                   R"json5({
  id: "builtin/dark",
  name: "Invalid Builtin",
  baseThemeId: "builtin/dark",
  colors: {
    "app.accent": "#445566",
  },
})json5"),
                 L"Failed to write invalid SettingsStore theme fixture.") &&
         ok;

    std::vector<Common::Settings::ThemeDefinition> themes;
    const HRESULT hr = Common::Settings::LoadThemeDefinitionsFromDirectory(tempRoot, themes);
    ok               = Require(SUCCEEDED(hr), L"SettingsStore theme loader should keep valid themes when one file is invalid.") && ok;
    ok               = Require(themes.size() == 1u, L"SettingsStore theme loader should reject ids rejected by ParseThemeDefinitionJson5.") && ok;
    if (themes.size() == 1u)
    {
        ok = Require(themes.front().id == L"user/settings-good", L"SettingsStore theme loader should preserve the valid theme.") && ok;
    }

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestRcStringTableParser()
{
    constexpr std::wstring_view rcText = LR"rc(
// Leading comments are ignored.
STRINGTABLE
BEGIN
    IDS_HELLO "Hello"
    IDS_QUOTED "He said ""Hello"""
END

STRINGTABLE
BEGIN
    IDS_FORMAT "Value {0:08X}"
    IDS_HELLO "Duplicate"
END
)rc";

    RedConfigure::Localization::RcParseResult parsed;
    const HRESULT hr = RedConfigure::Localization::ParseRcStringTables(rcText, parsed);
    bool ok          = true;
    ok               = Require(SUCCEEDED(hr), L"STRINGTABLE parser should accept the fixture.") && ok;
    ok               = Require(parsed.strings.size() == 4u, L"STRINGTABLE parser should read entries from multiple blocks.") && ok;
    ok               = Require(parsed.strings[1].text == L"He said \"Hello\"", L"STRINGTABLE parser should decode doubled quotes.") && ok;
    ok               = Require(parsed.strings[3].duplicate, L"STRINGTABLE parser should mark duplicate IDs.") && ok;
    ok               = Require(! parsed.errors.empty(), L"STRINGTABLE parser should report duplicate IDs.") && ok;
    return ok;
}

[[nodiscard]] bool TestRcMenuParser()
{
    constexpr std::wstring_view rcText = LR"rc(
IDR_MAIN MENU
BEGIN
    POPUP "&File"
    BEGIN
        MENUITEM "E&xit", IDM_EXIT
        POPUP "&Recent"
        BEGIN
            MENUITEM "&One", IDM_RECENT_ONE
        END
        MENUITEM SEPARATOR
    END
END
)rc";

    RedConfigure::Localization::RcParseResult parsed;
    const HRESULT hr = RedConfigure::Localization::ParseRcStringTables(rcText, parsed);
    bool ok          = true;
    ok               = Require(SUCCEEDED(hr), L"MENU parser should accept the fixture.") && ok;
    ok               = Require(parsed.localizableEntries.size() == 4u, L"MENU parser should expose popup and menu item captions.") && ok;
    if (parsed.localizableEntries.size() == 4u)
    {
        ok = Require(parsed.localizableEntries[0].kind == RedConfigure::Localization::RcLocalizableKind::MenuPopup,
                     L"MENU parser should classify POPUP captions.") &&
             ok;
        ok = Require(parsed.localizableEntries[0].ownerId == L"IDR_MAIN", L"MENU parser should preserve the menu resource id.") && ok;
        ok = Require(parsed.localizableEntries[0].text == L"&File", L"MENU parser should preserve popup text.") && ok;
        ok = Require(parsed.localizableEntries[1].kind == RedConfigure::Localization::RcLocalizableKind::MenuItem,
                     L"MENU parser should classify MENUITEM captions.") &&
             ok;
        ok = Require(parsed.localizableEntries[1].id == L"IDM_EXIT", L"MENU parser should preserve menu item ids.") && ok;
        ok = Require(parsed.localizableEntries[2].kind == RedConfigure::Localization::RcLocalizableKind::MenuPopup,
                     L"MENU parser should classify nested POPUP captions.") &&
             ok;
        ok = Require(parsed.localizableEntries[2].text == L"&Recent", L"MENU parser should parse nested popup text.") && ok;
        ok = Require(parsed.localizableEntries[3].id == L"IDM_RECENT_ONE", L"MENU parser should parse nested menu items.") && ok;
    }

    return ok;
}

[[nodiscard]] bool TestRcDialogParser()
{
    constexpr std::wstring_view rcText = LR"rc(
IDD_SAMPLE DIALOGEX 0, 0, 220, 90
CAPTION "Sample dialog"
BEGIN
    LTEXT           "Name:", IDC_STATIC, 7, 7, 48, 8
    PUSHBUTTON      "&OK", IDOK, 80, 68, 50, 14
    CONTROL         "&Enabled", IDC_ENABLED, "Button", BS_AUTOCHECKBOX | WS_TABSTOP, 7, 24, 80, 10
END
)rc";

    RedConfigure::Localization::RcParseResult parsed;
    const HRESULT hr = RedConfigure::Localization::ParseRcStringTables(rcText, parsed);
    bool ok          = true;
    ok               = Require(SUCCEEDED(hr), L"DIALOG parser should accept the fixture.") && ok;
    ok               = Require(parsed.localizableEntries.size() == 4u, L"DIALOG parser should expose captions and static/control text.") && ok;
    if (parsed.localizableEntries.size() == 4u)
    {
        ok = Require(parsed.localizableEntries[0].kind == RedConfigure::Localization::RcLocalizableKind::DialogCaption,
                     L"DIALOG parser should classify dialog captions.") &&
             ok;
        ok = Require(parsed.localizableEntries[0].ownerId == L"IDD_SAMPLE", L"DIALOG parser should preserve the dialog resource id.") && ok;
        ok = Require(parsed.localizableEntries[0].text == L"Sample dialog", L"DIALOG parser should parse captions.") && ok;
        ok = Require(parsed.localizableEntries[1].kind == RedConfigure::Localization::RcLocalizableKind::DialogControl,
                     L"DIALOG parser should classify control text.") &&
             ok;
        ok = Require(parsed.localizableEntries[1].id == L"IDC_STATIC", L"DIALOG parser should preserve control ids.") && ok;
        ok = Require(parsed.localizableEntries[2].id == L"IDOK", L"DIALOG parser should parse button ids.") && ok;
        ok = Require(parsed.localizableEntries[3].text == L"&Enabled", L"DIALOG parser should parse CONTROL text.") && ok;
    }

    return ok;
}

[[nodiscard]] bool TestPlaceholderValidation()
{
    using RedConfigure::Localization::ValidatePlaceholders;

    bool ok = true;
    ok      = Require(ValidatePlaceholders(L"Value {0:08X}", L"Valeur {0:08X}").status == RedConfigure::Localization::PlaceholderStatus::Ok,
                 L"Indexed placeholder validation should accept matching placeholders.") &&
         ok;
    ok = Require(ValidatePlaceholders(L"Value {0}", L"Valeur {}").status == RedConfigure::Localization::PlaceholderStatus::BarePlaceholder,
                 L"Bare placeholders should be rejected.") &&
         ok;
    ok = Require(ValidatePlaceholders(L"Value {0:08X}", L"Valeur {:08X}").status == RedConfigure::Localization::PlaceholderStatus::UnindexedFormatSpec,
                 L"Unindexed format specs should be rejected.") &&
         ok;
    ok = Require(ValidatePlaceholders(L"Value {0}", L"Valeur").status == RedConfigure::Localization::PlaceholderStatus::PlaceholderMismatch,
                 L"Missing target placeholders should be rejected.") &&
         ok;
    ok = Require(ValidatePlaceholders(L"Values {0} {0}", L"Valeurs {0}").status ==
                     RedConfigure::Localization::PlaceholderStatus::PlaceholderMismatch,
                 L"Duplicated source placeholders must not collapse to one target placeholder.") &&
         ok;
    ok = Require(ValidatePlaceholders(L"Value {0}", L"Valeur {0} {0}").status ==
                     RedConfigure::Localization::PlaceholderStatus::PlaceholderMismatch,
                 L"Duplicated target placeholders must not collapse to one source placeholder.") &&
         ok;
    ok = Require(ValidatePlaceholders(L"Value {0}", L"Valeur %s").status == RedConfigure::Localization::PlaceholderStatus::PrintfPlaceholder,
                 L"Printf placeholders should be rejected.") &&
         ok;
    return ok;
}

[[nodiscard]] bool TestTranslationViewSearchFilterAndSort()
{
    using RedConfigure::Localization::PlaceholderStatus;
    bool ok = true;

    std::vector<RedConfigure::TranslationEntry> rows;
    rows.push_back(RedConfigure::TranslationEntry{
        .id = L"IDS_SAVE",
        .sourceText = L"Save file",
        .targetText = L"Enregistrer",
        .validation = RedConfigure::Localization::PlaceholderValidationResult{.status = PlaceholderStatus::Ok}});
    rows.push_back(RedConfigure::TranslationEntry{
        .id = L"IDS_OPEN",
        .sourceText = L"Open file",
        .targetText = L"Ouvrir",
        .validation = RedConfigure::Localization::PlaceholderValidationResult{.status = PlaceholderStatus::Ok}});
    rows.push_back(RedConfigure::TranslationEntry{
        .id = L"IDS_DELETE",
        .sourceText = L"Delete {0}",
        .targetText = L"Supprimer",
        .validation = RedConfigure::Localization::PlaceholderValidationResult{.status = PlaceholderStatus::PlaceholderMismatch}});

    RedConfigure::LocalizationViewOptions options;
    options.searchText = L"file";
    std::vector<size_t> view = RedConfigure::BuildTranslationView(rows, options);
    ok = Require(view == std::vector<size_t>{0u, 1u}, L"Translation view search should match source and preserve source order.") && ok;

    options.searchText = {};
    options.idFilterText = L"delete";
    options.statusFilter = RedConfigure::LocalizationStatusFilter::Problems;
    view = RedConfigure::BuildTranslationView(rows, options);
    ok = Require(view == std::vector<size_t>{2u}, L"Translation view filters should combine ID and status filters.") && ok;

    options.idFilterText = {};
    options.statusFilter = RedConfigure::LocalizationStatusFilter::All;
    options.sortColumn = RedConfigure::LocalizationViewColumn::Target;
    options.sortDirection = RedConfigure::LocalizationSortDirection::Ascending;
    view = RedConfigure::BuildTranslationView(rows, options);
    ok = Require(view == std::vector<size_t>{0u, 1u, 2u}, L"Translation view target sort should order by target text ascending.") && ok;

    options.sortDirection = RedConfigure::LocalizationSortDirection::Descending;
    view = RedConfigure::BuildTranslationView(rows, options);
    ok = Require(view == std::vector<size_t>{2u, 1u, 0u}, L"Translation view target sort should reverse when descending.") && ok;

    return ok;
}

[[nodiscard]] bool TestRcWriterAndMerge()
{
    RedConfigure::Localization::RcParseResult source;
    RedConfigure::Localization::RcParseResult target;
    bool ok = true;

    ok = Require(SUCCEEDED(RedConfigure::Localization::ParseRcStringTables(LR"rc(
STRINGTABLE
BEGIN
    IDS_BETA "Beta {0}"
    IDS_ALPHA "Alpha"
END
)rc",
                                                                     source)),
                 L"Source STRINGTABLE fixture should parse.") &&
         ok;
    ok = Require(SUCCEEDED(RedConfigure::Localization::ParseRcStringTables(LR"rc(
STRINGTABLE
BEGIN
    IDS_BETA "Beta FR {0}"
END
)rc",
                                                                     target)),
                 L"Target STRINGTABLE fixture should parse.") &&
         ok;

    const auto merged = RedConfigure::Localization::MergeStringTables(source.strings, target.strings);
    ok               = Require(merged.size() == 2u, L"Merged resource model should contain all source strings.") && ok;
    ok               = Require(merged[1].targetText == L"Beta FR {0}", L"Merged resource model should keep target translations by ID.") && ok;

    const std::wstring output = RedConfigure::Localization::BuildSatelliteRcStringTable(L"resource.h", L"fr-FR", merged);
    ok                       = Require(output.find(L"#include \"resource.h\"") != std::wstring::npos, L"Satellite writer should include resource.h.") && ok;
    ok                       = Require(output.find(L"IDS_ALPHA") < output.find(L"IDS_BETA"), L"Satellite writer should sort entries by ID.") && ok;
    ok = Require(output.find(L"\"Beta FR {0}\"") != std::wstring::npos, L"Satellite writer should keep positional placeholders unchanged.") && ok;
    return ok;
}

[[nodiscard]] bool TestThemeCatalogLoadsThemeFiles()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(ec) / L"RedConfigureThemeCatalogTest";
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for theme catalog tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    bool ok = true;
    ok      = Require(WriteTestTextFile(tempRoot / L"Good.theme.json5",
                                   R"json5({
  id: "user/catalog-good",
  name: "Catalog Good",
  baseThemeId: "builtin/dark",
  colors: {
    "app.accent": "#336699",
  },
})json5"),
                 L"Failed to write valid theme catalog fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"Bad.theme.json5", "{ id: 'user/bad' }\n"), L"Failed to write invalid theme catalog fixture.") && ok;

    std::vector<RedConfigure::Workspace::ThemeFile> files;
    files.push_back({.path = tempRoot / L"Good.theme.json5"});
    files.push_back({.path = tempRoot / L"Bad.theme.json5"});

    RedConfigure::Themes::ThemeCatalog catalog;
    const HRESULT hr = RedConfigure::Themes::LoadThemeCatalog(files, catalog);
    ok               = Require(SUCCEEDED(hr), L"Theme catalog load should return a partial result when one file is invalid.") && ok;
    ok               = Require(catalog.themes.size() == 1u, L"Theme catalog should include the one valid theme.") && ok;
    ok               = Require(catalog.errors.size() == 1u, L"Theme catalog should surface invalid theme file errors.") && ok;
    if (! catalog.themes.empty())
    {
        ok = Require(catalog.themes.front().definition.id == L"user/catalog-good", L"Theme catalog should preserve parsed theme ids.") && ok;
    }

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestThemePreviewModelKeepsLastValidColor()
{
    Common::Settings::ThemeDefinition theme;
    theme.id          = L"user/preview";
    theme.name        = L"Preview";
    theme.baseThemeId = L"builtin/dark";
    theme.colors.emplace(L"folderView.itemBackgroundSelected", 0xFF204060u);

    RedConfigure::Themes::ThemePreviewModel model;
    model.SetTheme(theme);

    bool ok = true;
    ok      = Require(model.GetEffectiveColor(L"folderView.itemBackgroundSelected").value_or(0u) == 0xFF204060u,
                 L"Theme preview should expose the initial override color.") &&
         ok;
    ok = Require(model.TryEditOverride(L"folderView.itemBackgroundSelected", L"#336699"), L"Theme preview should accept valid color edits.") && ok;
    ok = Require(model.GetEffectiveColor(L"folderView.itemBackgroundSelected").value_or(0u) == 0xFF336699u,
                 L"Theme preview should update immediately after valid color edits.") &&
         ok;
    ok = Require(! model.TryEditOverride(L"folderView.itemBackgroundSelected", L"not-a-color"), L"Theme preview should reject invalid color text.") && ok;
    ok = Require(model.GetEffectiveColor(L"folderView.itemBackgroundSelected").value_or(0u) == 0xFF336699u,
                 L"Theme preview should keep the last valid color after invalid text.") &&
         ok;
    return ok;
}

[[nodiscard]] bool TestThemePreviewModelResolvesColorExpressions()
{
    Common::Settings::ThemeDefinition theme;
    theme.id          = L"user/expressions";
    theme.name        = L"Expressions";
    theme.baseThemeId = L"builtin/dark";
    theme.colors.emplace(L"app.accent", 0xFF808080u);
    theme.colors.emplace(L"menu.background", 0xFF202020u);

    RedConfigure::Themes::ThemePreviewModel model;
    model.SetTheme(theme);

    bool ok = true;
    ok = Require(model.TryEditOverride(L"folderView.itemBackgroundSelected", L"darken(app.accent,25%)"),
                 L"Theme preview should accept darken expressions.") &&
         ok;
    ok = Require(model.GetEffectiveColor(L"folderView.itemBackgroundSelected").value_or(0u) == 0xFF606060u,
                 L"Darken expressions should derive the expected color.") &&
         ok;

    ok = Require(model.TryEditOverride(L"folderView.background", L"blend(menu.background,app.accent,50%)"),
                 L"Theme preview should accept blend expressions.") &&
         ok;
    ok = Require(model.GetEffectiveColor(L"folderView.background").value_or(0u) == 0xFF505050u,
                 L"Blend expressions should derive the expected color.") &&
         ok;

    ok = Require(model.TryEditOverride(L"window.background", L"ref(folderView.background)"),
                 L"Theme preview should accept reference expressions.") &&
         ok;
    ok = Require(model.GetEffectiveColor(L"window.background").value_or(0u) == 0xFF505050u,
                 L"Reference expressions should reuse the referenced effective color.") &&
         ok;

    ok = Require(! model.TryEditOverride(L"app.accent", L"ref(folderView.itemBackgroundSelected)"),
                 L"Theme preview should reject color expression cycles.") &&
         ok;
    ok = Require(model.GetEffectiveColor(L"app.accent").value_or(0u) == 0xFF808080u,
                 L"Rejected expression cycles should keep the previous valid color.") &&
         ok;
    return ok;
}

[[nodiscard]] bool TestThemeColorSuggestionsGuideExpressionEditing()
{
    const std::vector<std::wstring> suggestions =
        RedConfigure::BuildThemeColorSuggestions(L"folderView.itemBackgroundSelected", L"menu.background", 0xFF2ECC71u);

    const auto contains = [&suggestions](std::wstring_view expected) noexcept
    {
        return std::find(suggestions.begin(), suggestions.end(), expected) != suggestions.end();
    };

    bool ok = true;
    ok = Require(contains(L"#2ECC71"), L"Theme color suggestions should include the current direct color.") && ok;
    ok = Require(contains(L"ref(app.accent)"), L"Theme color suggestions should include an accent reference.") && ok;
    ok = Require(contains(L"darken(app.accent,20%)"), L"Theme color suggestions should include a darken expression template.") && ok;
    ok = Require(contains(L"blend(menu.background,app.accent,16%)"),
                 L"Theme color suggestions should include a blend expression using the previously selected color.") &&
         ok;
    ok = Require(contains(L"ref(menu.background)"), L"Theme color suggestions should include the previous color reference.") && ok;
    return ok;
}

[[nodiscard]] bool TestThemeColorKeyFilterNarrowsKeysCaseInsensitively()
{
    const std::vector<std::wstring> keys = {
        L"app.accent",
        L"menu.background",
        L"menu.selectionBackground",
        L"folderView.background",
        L"dialog.text",
    };

    bool ok = true;
    std::vector<std::wstring> filtered = RedConfigure::FilterThemeColorKeys(keys, L"MENU");
    ok = Require(filtered == std::vector<std::wstring>{L"menu.background", L"menu.selectionBackground"},
                 L"Theme color key filter should match key groups case-insensitively.") &&
         ok;

    filtered = RedConfigure::FilterThemeColorKeys(keys, L"background");
    ok = Require(filtered == std::vector<std::wstring>{L"menu.background", L"menu.selectionBackground", L"folderView.background"},
                 L"Theme color key filter should match substrings anywhere in the key.") &&
         ok;

    filtered = RedConfigure::FilterThemeColorKeys(keys, L"missing");
    ok = Require(filtered.empty(), L"Theme color key filter should produce an empty list when no keys match.") && ok;
    return ok;
}

[[nodiscard]] bool TestThemePreviewHitSelectionUsesSmallestRegionAndCycles()
{
    const std::vector<RedConfigure::ThemePreviewHitCandidate> candidates = {
        RedConfigure::ThemePreviewHitCandidate{.key = L"navigation.background", .left = 0.0f, .top = 0.0f, .right = 200.0f, .bottom = 200.0f},
        RedConfigure::ThemePreviewHitCandidate{.key = L"app.accent", .left = 0.0f, .top = 0.0f, .right = 6.0f, .bottom = 200.0f},
        RedConfigure::ThemePreviewHitCandidate{.key = L"menu.selectionBackground", .left = 60.0f, .top = 10.0f, .right = 120.0f, .bottom = 40.0f},
        RedConfigure::ThemePreviewHitCandidate{.key = L"menu.background", .left = 40.0f, .top = 0.0f, .right = 240.0f, .bottom = 60.0f},
    };

    bool ok = true;
    ok = Require(RedConfigure::SelectThemePreviewHitKey(candidates, 3.0f, 50.0f, {}) == L"app.accent",
                 L"Theme preview hit selection should prefer the smallest visible region.") &&
         ok;
    ok = Require(RedConfigure::SelectThemePreviewHitKey(candidates, 3.0f, 50.0f, L"app.accent") == L"navigation.background",
                 L"Theme preview repeated clicks should cycle to the containing region.") &&
         ok;
    ok = Require(RedConfigure::SelectThemePreviewHitKey(candidates, 80.0f, 20.0f, {}) == L"menu.selectionBackground",
                 L"Theme preview hit selection should prefer nested menu selection over menu background.") &&
         ok;
    ok = Require(RedConfigure::SelectThemePreviewHitKey(candidates, 500.0f, 500.0f, {}).empty(),
                 L"Theme preview hit selection should return no key outside all regions.") &&
         ok;
    return ok;
}

[[nodiscard]] bool TestRedConfigureSessionExportsFirstUsableFiles()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(ec) / L"RedConfigureSessionTest";
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for session tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    bool ok = true;
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"App.vcxproj",
                                   R"xml(<?xml version="1.0" encoding="utf-8"?>
<Project>
  <ItemGroup>
    <ResourceCompile Include="App.rc" />
  </ItemGroup>
</Project>)xml"),
                 L"Failed to write session project fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"App" / L"App.rc",
                                   R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Hello {0}"
END
IDR_APP MENU
BEGIN
    POPUP "&File"
    BEGIN
        MENUITEM "E&xit", IDM_EXIT
    END
END
IDD_APP DIALOGEX 0, 0, 160, 80
CAPTION "App dialog"
BEGIN
    LTEXT "Name:", IDC_STATIC, 7, 7, 48, 8
END
)rc"),
                 L"Failed to write session source rc fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"App" / L"Lang" / L"fr-FR" / L"App-fr-FR.rc",
                                   "#include \"resource.h\"\nSTRINGTABLE\nBEGIN\n    IDS_HELLO \"Bonjour {0}\"\nEND\n"),
                 L"Failed to write session target rc fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"Zeta" / L"Zeta.vcxproj",
                                   R"xml(<?xml version="1.0" encoding="utf-8"?>
<Project>
  <ItemGroup>
    <ResourceCompile Include="Zeta.rc" />
  </ItemGroup>
</Project>)xml"),
                 L"Failed to write second session project fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"Zeta" / L"Zeta.rc", "STRINGTABLE\nBEGIN\n    IDS_ZETA \"Zeta\"\nEND\n"),
                 L"Failed to write second source rc fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"Themes" / L"Session.theme.json5",
                                   R"json5({
  id: "user/session",
  name: "Session",
  baseThemeId: "builtin/dark",
  colors: {
    "app.accent": "#336699",
  },
})json5"),
                 L"Failed to write session theme fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"Themes" / L"Zed.theme.json5",
                                   R"json5({
  id: "user/zed",
  name: "Zed",
  baseThemeId: "builtin/light",
  colors: {
    "app.accent": "#AA5500",
  },
})json5"),
                 L"Failed to write second session theme fixture.") &&
         ok;

    RedConfigure::RedConfigureSession session;
    ok = Require(SUCCEEDED(session.LoadWorkspace(tempRoot, L"fr-FR")), L"Session should load a fixture workspace.") && ok;
    ok = Require(session.GetTranslations().size() == 1u, L"Session should expose merged translations.") && ok;
    ok = Require(session.GetInventoryEntries().size() == 5u, L"Session inventory should include strings, menu captions, and dialog text.") && ok;
    if (! session.GetTranslations().empty())
    {
        const auto& first = session.GetTranslations().front();
        ok               = Require(first.id == L"IDS_HELLO", L"Session should preserve string ids.") && ok;
        ok               = Require(first.targetText == L"Bonjour {0}", L"Session should merge target translations.") && ok;
    }
    ok = Require(SUCCEEDED(session.SetActiveResourceOwner(1u)), L"Session should switch active resource owners.") && ok;
    ok = Require(! session.GetTranslations().empty() && session.GetTranslations().front().id == L"IDS_ZETA",
                 L"Session should reload translations when the owner changes.") &&
         ok;
    ok = Require(session.SetActiveTheme(1u), L"Session should switch active themes.") && ok;
    ok = Require(session.GetThemePreviewModel().GetTheme().id == L"user/zed", L"Session should update the preview theme when switched.") && ok;
    ok = Require(SUCCEEDED(session.SetActiveResourceOwner(0u)), L"Session should switch back to the first resource owner.") && ok;

    ok = Require(session.UpdateTranslation(0u, L"Salut {0}"), L"Session should accept a valid translation edit.") && ok;
    ok = Require(! session.UpdateThemeColor(L"app.accent", L"not-a-color"), L"Session should reject invalid theme colors.") && ok;
    ok = Require(session.UpdateThemeColor(L"app.accent", L"#123456"), L"Session should accept valid theme colors.") && ok;
    ok = Require(session.UpdateThemeColor(L"folderView.itemBackgroundSelected", L"darken(app.accent,50%)"),
                 L"Session should accept theme expression color edits.") &&
         ok;

    std::wstring rcPreview;
    std::string themePreview;
    ok = Require(SUCCEEDED(session.BuildLocalizationExportText(rcPreview)), L"Session should build a localization export preview.") && ok;
    ok = Require(SUCCEEDED(session.BuildThemeExportText(themePreview)), L"Session should build a theme export preview.") && ok;
    ok = Require(rcPreview.find(L"Salut {0}") != std::wstring::npos, L"Localization export preview should contain edited translations.") && ok;
    ok = Require(themePreview.find("\"app.accent\": \"#123456\"") != std::string::npos,
                 L"Theme export preview should contain edited colors.") &&
         ok;
    ok = Require(themePreview.find("\"folderView.itemBackgroundSelected\": \"#091A2B\"") != std::string::npos,
                 L"Theme export preview should flatten expression-authored colors.") &&
         ok;

    const std::filesystem::path rcPath    = tempRoot / L"Out" / L"App-fr-FR.rc";
    const std::filesystem::path themePath = tempRoot / L"Out" / L"Session.theme.json5";
    ok = Require(SUCCEEDED(session.ExportLocalization(rcPath)), L"Session should export a satellite rc file.") && ok;
    ok = Require(SUCCEEDED(session.ExportTheme(themePath)), L"Session should export a theme json5 file.") && ok;
    ok = Require(std::filesystem::exists(rcPath, ec), L"Exported rc file should exist.") && ok;
    ok = Require(std::filesystem::exists(themePath, ec), L"Exported theme file should exist.") && ok;

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestRedConfigureSessionReadsBomlessUtf16LeRcFiles()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(ec) / L"RedConfigureBomlessUtf16LeSessionTest";
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for BOM-less UTF-16 session tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    bool ok = true;
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"App.vcxproj",
                                   R"xml(<?xml version="1.0" encoding="utf-8"?>
<Project>
  <ItemGroup>
    <ResourceCompile Include="App.rc" />
  </ItemGroup>
</Project>)xml"),
                 L"Failed to write BOM-less session project fixture.") &&
         ok;
    ok = Require(WriteTestUtf16LeTextFile(tempRoot / L"App" / L"App.rc",
                                          LR"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Héllo {0}"
END
)rc",
                                          false),
                 L"Failed to write BOM-less UTF-16 LE source rc fixture.") &&
         ok;
    ok = Require(WriteTestUtf16LeTextFile(tempRoot / L"App" / L"Lang" / L"fr-FR" / L"App-fr-FR.rc",
                                          LR"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Salut {0}"
END
)rc",
                                          false),
                 L"Failed to write BOM-less UTF-16 LE target rc fixture.") &&
         ok;

    RedConfigure::RedConfigureSession session;
    ok = Require(SUCCEEDED(session.LoadWorkspace(tempRoot, L"fr-FR")), L"Session should load BOM-less UTF-16 LE rc files.") && ok;
    ok = Require(session.GetTranslations().size() == 1u, L"Session should parse one string from BOM-less UTF-16 LE rc files.") && ok;
    if (session.GetTranslations().size() == 1u)
    {
        ok = Require(session.GetTranslations().front().sourceText == L"Héllo {0}",
                     L"Session should preserve non-ASCII source text from BOM-less UTF-16 LE.") &&
             ok;
        ok = Require(session.GetTranslations().front().targetText == L"Salut {0}",
                     L"Session should preserve target text from BOM-less UTF-16 LE satellite resources.") &&
             ok;
    }

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}
} // namespace

int wmain()
{
    if (! TestPageDefinitions())
    {
        return 1;
    }
    if (! TestResolveWorkspaceRootForLaunchPath())
    {
        return 1;
    }
    if (! TestWorkspaceDiscoveryFindsResourcesAndThemes())
    {
        return 1;
    }
    if (! TestWorkspaceDiscoveryReportsScanErrors())
    {
        return 1;
    }
    if (! TestThemeDefinitionJson5Parsing())
    {
        return 1;
    }
    if (! TestThemeDefinitionRejectsInvalidInput())
    {
        return 1;
    }
    if (! TestThemeDefinitionJson5Export())
    {
        return 1;
    }
    if (! TestSettingsStoreThemeDirectoryUsesSharedParser())
    {
        return 1;
    }
    if (! TestRcStringTableParser())
    {
        return 1;
    }
    if (! TestRcMenuParser())
    {
        return 1;
    }
    if (! TestRcDialogParser())
    {
        return 1;
    }
    if (! TestPlaceholderValidation())
    {
        return 1;
    }
    if (! TestTranslationViewSearchFilterAndSort())
    {
        return 1;
    }
    if (! TestRcWriterAndMerge())
    {
        return 1;
    }
    if (! TestThemeCatalogLoadsThemeFiles())
    {
        return 1;
    }
    if (! TestThemePreviewModelKeepsLastValidColor())
    {
        return 1;
    }
    if (! TestThemePreviewModelResolvesColorExpressions())
    {
        return 1;
    }
    if (! TestThemeColorSuggestionsGuideExpressionEditing())
    {
        return 1;
    }
    if (! TestThemeColorKeyFilterNarrowsKeysCaseInsensitively())
    {
        return 1;
    }
    if (! TestThemePreviewHitSelectionUsesSmallestRegionAndCycles())
    {
        return 1;
    }
    if (! TestRedConfigureSessionExportsFirstUsableFiles())
    {
        return 1;
    }
    if (! TestRedConfigureSessionReadsBomlessUtf16LeRcFiles())
    {
        return 1;
    }

    std::wcout << L"RedConfigureTests passed.\n";
    return 0;
}
