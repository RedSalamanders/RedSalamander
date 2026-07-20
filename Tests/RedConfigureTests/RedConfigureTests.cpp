#include "Localization/PlaceholderValidation.h"
#include "Localization/RcParser.h"
#include "Localization/RcWriter.h"
#include "Helpers.h"
#include "RedConfigureApp.h"
#include "RedConfigureBinaryFile.h"
#include "RedConfigureGridModels.h"
#include "RedConfigurePagePresenters.h"
#include "RedConfigureRoot.h"
#include "RedConfigureSession.h"
#include "RedConfigureWorkflow.h"
#include "RedConfigureSplashScreen.h"
#include "SettingsStore.h"
#include "ThemeDefinitionIo.h"
#include "Themes/ThemeCatalog.h"
#include "Themes/ThemePreviewModel.h"
#include "Workspace/WorkspaceDiscovery.h"
#include "TestSupport/TestSupport.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <clocale>
#include <cmath>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <format>
#include <iostream>
#include <optional>
#include <span>
#include <stop_token>
#include <string>
#include <string_view>
#include <system_error>
#include <vector>

#include <Windows.h>

#pragma warning(push)
// WIL headers: deleted copy/move and unused inline helpers
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace
{
constexpr std::wstring_view kRedConfigureHarnessSegment{L"redconfigure"};

[[nodiscard]] bool Require(bool condition, std::wstring_view message)
{
    if (! condition)
    {
        std::wcerr << message << L'\n';
        return false;
    }

    return true;
}

[[nodiscard]] std::wstring GetEnvironmentString(std::wstring_view name)
{
    return RedSalamander::TestSupport::GetEnvironmentString(name);
}

[[nodiscard]] std::filesystem::path AcquireRedConfigureTestSandbox(std::wstring_view caseName, std::error_code& ec) noexcept
{
    return RedSalamander::TestSupport::AcquireTestDirectory({.harnessSegment      = kRedConfigureHarnessSegment,
                                                             .leafSegment         = caseName,
                                                             .fallbackRunIdPrefix = L"redconfigure",
                                                             .cleanExisting       = false},
                                                            ec);
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

void RemoveSettingsFiles(std::wstring_view appId) noexcept
{
    std::error_code ec;
    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(appId);
    if (! settingsPath.empty())
    {
        std::filesystem::remove(settingsPath, ec);
    }

    const std::filesystem::path schemaPath = Common::Settings::GetSettingsSchemaPath(appId);
    if (! schemaPath.empty())
    {
        std::filesystem::remove(schemaPath, ec);
    }
}

[[nodiscard]] bool WriteTestBinaryFile(const std::filesystem::path& path, std::span<const uint8_t> bytes)
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

    if (! bytes.empty())
    {
        output.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
    }
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
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureWorkspaceRootResolveTest", ec);
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
    ok                                   = Require(resolved == tempRoot, L"Launch paths inside .build should resolve back to the repository root.") && ok;

    const std::filesystem::path unknownRoot     = std::filesystem::path(L"RedConfigureStandaloneRootResolveTest");
    const std::filesystem::path unknownResolved = RedConfigure::ResolveWorkspaceRootForLaunchPath(unknownRoot);
    ok                                          = Require(unknownResolved == unknownRoot, L"Paths without repository markers should be preserved.") && ok;

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestBoundedBinaryFileReader()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureBinaryFileReaderTest", ec);
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for binary-reader tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    ec.clear();

    const std::filesystem::path emptyPath   = tempRoot / L"empty.bin";
    const std::filesystem::path unicodePath = tempRoot / L"données-測試.bin";
    const std::array<uint8_t, 4u> bytes{0x00u, 0x7Fu, 0x80u, 0xFFu};
    bool ok = Require(WriteTestBinaryFile(emptyPath, {}), L"Failed to write empty binary-reader fixture.");
    ok = Require(WriteTestBinaryFile(unicodePath, bytes), L"Failed to write Unicode binary-reader fixture.") && ok;

    std::vector<uint8_t> actual{0xAAu};
    HRESULT hr = RedConfigure::ReadBinaryFile(emptyPath, actual, 0u);
    ok = Require(hr == S_OK && actual.empty(), L"The binary reader must accept an empty file at a zero-byte bound.") && ok;

    hr = RedConfigure::ReadBinaryFile(unicodePath, actual, bytes.size());
    ok = Require(hr == S_OK && std::ranges::equal(actual, bytes), L"The binary reader must preserve bytes at the exact bound and on Unicode paths.") && ok;

    actual.assign(1u, 0xAAu);
    hr = RedConfigure::ReadBinaryFile(unicodePath, actual, bytes.size() - 1u);
    ok = Require(hr == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE) && actual.empty(),
                 L"The binary reader must reject an oversized file without publishing partial bytes.") &&
         ok;

    actual.assign(1u, 0xAAu);
    hr = RedConfigure::ReadBinaryFile(tempRoot / L"missing.bin", actual, bytes.size());
    ok = Require(hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && actual.empty(),
                 L"The binary reader must preserve the missing-file HRESULT and clear prior output.") &&
         ok;

    wil::unique_handle locked(::CreateFileW(unicodePath.c_str(), GENERIC_READ, 0u, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    ok = Require(static_cast<bool>(locked), L"Failed to open the binary-reader sharing fixture exclusively.") && ok;
    if (locked)
    {
        actual.assign(1u, 0xAAu);
        hr = RedConfigure::ReadBinaryFile(unicodePath, actual, bytes.size());
        ok = Require(hr == HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) && actual.empty(),
                     L"The binary reader must return the locked-file error without publishing bytes.") &&
             ok;
    }

    locked.reset();
    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestWorkspaceDiscoveryFindsResourcesAndThemes()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureWorkspaceDiscoveryTest", ec);
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
    ok = Require(WriteTestTextFile(tempRoot / L"RedSalamander" / L"RedSalamander.rc", "STRINGTABLE\nBEGIN\nEND\n"), L"Failed to write app rc fixture.") && ok;
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
    ok =
        Require(WriteTestTextFile(tempRoot / L"Specs" / L"Themes" / L"Forest.theme.json5", "{ id: 'user/forest' }\n"), L"Failed to write theme fixture.") && ok;
    ok = Require(WriteTestTextFile(tempRoot / L".build" / L"x64" / L"Debug" / L"Ignored.vcxproj",
                                   R"xml(<Project><ItemGroup><ResourceCompile Include="Ignored.rc" /></ItemGroup></Project>)xml"),
                 L"Failed to write ignored build project fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L".build" / L"x64" / L"Debug" / L"Themes" / L"Ignored.theme.json5", "{}\n"),
                 L"Failed to write ignored build theme fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"vcpkg_installed" / L"x64-windows" / L"IgnoredDependency.vcxproj",
                                   R"xml(<Project><ItemGroup><ResourceCompile Include="IgnoredDependency.rc" /></ItemGroup></Project>)xml"),
                 L"Failed to write ignored dependency project fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"x64" / L"Debug" / L"IgnoredOutput.vcxproj",
                                   R"xml(<Project><ItemGroup><ResourceCompile Include="IgnoredOutput.rc" /></ItemGroup></Project>)xml"),
                 L"Failed to write ignored x64 output project fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L".git" / L"objects" / L"IgnoredGit.vcxproj",
                                   R"xml(<Project><ItemGroup><ResourceCompile Include="IgnoredGit.rc" /></ItemGroup></Project>)xml"),
                 L"Failed to write ignored git project fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L".vs" / L"IgnoredVs.vcxproj",
                                   R"xml(<Project><ItemGroup><ResourceCompile Include="IgnoredVs.rc" /></ItemGroup></Project>)xml"),
                 L"Failed to write ignored Visual Studio cache project fixture.") &&
         ok;
    ok = Require(WriteTestTextFile(tempRoot / L"Specs" / L"TestRuns" / L"IgnoredRun" / L"IgnoredRun.theme.json5", "{}\n"),
                 L"Failed to write ignored test-run theme fixture.") &&
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
    ok                   = Require(appOwner != nullptr, L"Workspace discovery should expose the RedSalamander owner.") && ok;
    if (appOwner)
    {
        ok = Require(appOwner->embeddedResourcePath.filename() == L"RedSalamander.rc", L"App owner should expose its embedded resource file.") && ok;
        ok = Require(appOwner->satelliteResourcePaths.size() == 1u, L"App owner should expose one satellite resource file.") && ok;
    }

    const auto* pluginOwner = ownerByName(L"ViewerText");
    ok                      = Require(pluginOwner != nullptr, L"Workspace discovery should expose the ViewerText owner.") && ok;
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
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureWorkspaceDiscoveryErrorTest", ec);
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
  "formatVersion": 2,
  "id": "user/forest-test",
  "name": "Forest Test",
  "baseThemeId": "builtin/dark",
  "palette": {
    "accent": "#2ECC71",
  },
  "colors": {
    "app.accent": "ref(palette.accent)",
    "folderView.itemBackgroundSelected": "alpha(palette.accent,50.196078%)",
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
    auto context = Common::Settings::MakeSystemThemeResolutionContext(true);
    Common::Settings::ResolvedThemeColors resolved;
    ok = Require(SUCCEEDED(Common::Settings::ResolveThemeDefinition(theme, context, resolved)), L"Parsed theme sources should resolve.") && ok;
    ok = Require(resolved.colors[L"app.accent"] == 0xFF2ECC71u, L"Palette reference did not resolve.") && ok;
    ok = Require(resolved.colors[L"folderView.itemBackgroundSelected"] == 0x802ECC71u, L"Alpha expression did not resolve.") && ok;

    return ok;
}

[[nodiscard]] bool TestThemeDefinitionRejectsInvalidInput()
{
    Common::Settings::ThemeDefinition theme;
    Common::Settings::ThemeDefinitionIoError error = Common::Settings::ThemeDefinitionIoError::None;
    bool ok                                        = true;

    const HRESULT emptyInput = Common::Settings::ParseThemeDefinitionJson5("", theme, &error, nullptr);
    ok                       = Require(emptyInput == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Empty theme input should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::EmptyInput, L"Empty theme input should report the empty-input error.") && ok;

    const HRESULT missingVersion = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"id":"user/legacy","name":"Legacy","baseThemeId":"builtin/dark","colors":{"app.accent":"#112233"}})json", theme, &error, nullptr);
    ok = Require(missingVersion == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) &&
                     error == Common::Settings::ThemeDefinitionIoError::MissingOrInvalidFormatVersion,
                 L"Legacy themes without formatVersion should be rejected deterministically.") && ok;

    const HRESULT versionOne = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"formatVersion":1,"id":"user/legacy","name":"Legacy","baseThemeId":"builtin/dark","colors":{}})json", theme, &error, nullptr);
    ok = Require(versionOne == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && error == Common::Settings::ThemeDefinitionIoError::UnsupportedFormatVersion,
                 L"Version 1 themes should be rejected deterministically.") && ok;

    const HRESULT missingId =
        Common::Settings::ParseThemeDefinitionJson5(R"json({"formatVersion":2,"name":"Bad","baseThemeId":"builtin/dark","colors":{}})json", theme, &error, nullptr);
    ok = Require(missingId == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Missing id should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::MissingOrInvalidId, L"Missing id should report the id error.") && ok;

    const HRESULT missingName =
        Common::Settings::ParseThemeDefinitionJson5(R"json({"formatVersion":2,"id":"user/bad","baseThemeId":"builtin/dark","colors":{}})json", theme, &error, nullptr);
    ok = Require(missingName == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Missing name should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::MissingOrInvalidName, L"Missing name should report the name error.") && ok;

    const HRESULT missingBaseTheme =
        Common::Settings::ParseThemeDefinitionJson5(R"json({"formatVersion":2,"id":"user/bad","name":"Bad","colors":{}})json", theme, &error, nullptr);
    ok = Require(missingBaseTheme == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Missing base theme should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::MissingOrInvalidBaseThemeId, L"Missing base theme should report the base-theme error.") &&
         ok;

    const HRESULT missingColors =
        Common::Settings::ParseThemeDefinitionJson5(R"json({"formatVersion":2,"id":"user/bad","name":"Bad","baseThemeId":"builtin/dark"})json", theme, &error, nullptr);
    ok = Require(missingColors == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Missing colors should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::ColorsMissingOrNotObject, L"Missing colors should report the colors error.") && ok;

    const HRESULT invalidId = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"formatVersion":2,"id":"builtin/dark","name":"Bad","baseThemeId":"builtin/dark","colors":{}})json", theme, &error, nullptr);
    ok = Require(invalidId == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Invalid user theme id should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::InvalidId, L"Invalid id should report the id error.") && ok;

    const HRESULT invalidBaseTheme =
        Common::Settings::ParseThemeDefinitionJson5(R"json({"formatVersion":2,"id":"user/bad","name":"Bad","baseThemeId":"user/base","colors":{}})json", theme, &error, nullptr);
    ok = Require(invalidBaseTheme == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Invalid base theme should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::InvalidBaseThemeId, L"Invalid base theme should report the base-theme error.") && ok;

    const HRESULT invalidColorKey = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"formatVersion":2,"id":"user/bad","name":"Bad","baseThemeId":"builtin/dark","colors":{"bad key":"#112233"}})json", theme, &error, nullptr);
    ok = Require(invalidColorKey == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Invalid color key should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::InvalidColorKey, L"Invalid color key should report the color-key error.") && ok;

    const HRESULT colorNotString = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"formatVersion":2,"id":"user/bad","name":"Bad","baseThemeId":"builtin/dark","colors":{"app.accent":12}})json", theme, &error, nullptr);
    ok = Require(colorNotString == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Non-string color should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::ColorValueNotString, L"Non-string color should report the color-string error.") && ok;

    const HRESULT invalidColor = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"formatVersion":2,"id":"user/bad","name":"Bad","baseThemeId":"builtin/dark","colors":{"app.accent":"red"}})json", theme, &error, nullptr);
    ok = Require(invalidColor == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Invalid color should fail with invalid data.") && ok;
    ok = Require(error == Common::Settings::ThemeDefinitionIoError::InvalidColorValue, L"Invalid color should report the color value error.") && ok;

    const HRESULT duplicatePalette = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"formatVersion":2,"id":"user/bad","name":"Bad","baseThemeId":"builtin/dark","palette":{"accent":"#112233","accent":"#445566"},"colors":{}})json",
        theme,
        &error,
        nullptr);
    ok = Require(duplicatePalette == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && error == Common::Settings::ThemeDefinitionIoError::DuplicatePaletteName,
                 L"Duplicate palette names should fail with a deterministic error.") && ok;

    const HRESULT invalidPaletteName = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"formatVersion":2,"id":"user/bad","name":"Bad","baseThemeId":"builtin/dark","palette":{"bad.name":"#112233"},"colors":{}})json",
        theme,
        &error,
        nullptr);
    ok = Require(invalidPaletteName == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && error == Common::Settings::ThemeDefinitionIoError::InvalidPaletteName,
                 L"Palette names containing a dot should fail with a deterministic error.") && ok;

    const HRESULT dynamicPalette = Common::Settings::ParseThemeDefinitionJson5(
        R"json({"formatVersion":2,"id":"user/bad","name":"Bad","baseThemeId":"builtin/dark","palette":{"accent":"systemAccent()"},"colors":{}})json",
        theme,
        &error,
        nullptr);
    ok = Require(dynamicPalette == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && error == Common::Settings::ThemeDefinitionIoError::InvalidColorValue,
                 L"Palette entries should reject event-time sources.") && ok;

    return ok;
}

[[nodiscard]] bool TestThemeDefinitionJson5Export()
{
    bool ok = true;
    Common::Settings::ThemeDefinition theme;
    theme.id          = L"user/export-test";
    theme.name        = L"Export Test";
    theme.baseThemeId = L"builtin/light";
    theme.palette.emplace(L"accent", 0xFF0055AAu);
    Common::Settings::ThemeColorSource accentReference;
    ok = Require(SUCCEEDED(Common::Settings::ParseThemeColorSource(L"ref(palette.accent)", accentReference)), L"Export test reference should parse.") && ok;
    theme.colors.emplace(L"app.accent", accentReference);
    theme.colors.emplace(L"folderView.background", 0xFFFFFFFFu);

    std::string json;
    const HRESULT hr = Common::Settings::BuildThemeDefinitionJson5(theme, json);
    ok               = Require(SUCCEEDED(hr), L"Expected theme export to succeed.") && ok;
    ok               = Require(json.find("// Named palette") != std::string::npos && json.find("// app") != std::string::npos,
                               L"Exported theme should include stable palette and color-group comments.") &&
                       ok;
    ok               = Require(json.find("\"id\": \"user/export-test\"") != std::string::npos, L"Exported theme id is missing.") && ok;
    ok = Require(json.find("\"app.accent\": \"ref(palette.accent)\"") < json.find("\"folderView.background\": \"#FFFFFF\""),
                 L"Exported colors should preserve sources and be sorted by key.") &&
         ok;

    Common::Settings::ThemeDefinition reparsed;
    Common::Settings::ThemeDefinitionIoError error = Common::Settings::ThemeDefinitionIoError::None;
    const HRESULT parseHr                          = Common::Settings::ParseThemeDefinitionJson5(json, reparsed, &error, nullptr);
    ok                                             = Require(SUCCEEDED(parseHr), L"Exported theme should parse again.") && ok;
    ok                                             = Require(reparsed.palette == theme.palette, L"Reparsed exported palette did not round trip.") && ok;
    ok                                             = Require(reparsed.colors == theme.colors, L"Reparsed exported colors did not round trip.") && ok;

    return ok;
}

[[nodiscard]] bool TestThemeExpressionLanguageAndRuntimeSources()
{
    bool ok = true;
    Common::Settings::ThemeDefinition theme;
    theme.id = L"user/expression-language";
    theme.name = L"Expression Language";
    theme.baseThemeId = L"builtin/dark";

    const auto addSource = [&](auto& destination, std::wstring key, std::wstring_view text)
    {
        Common::Settings::ThemeColorSource source;
        std::wstring message;
        const HRESULT hr = Common::Settings::ParseThemeColorSource(text, source, &message);
        ok = Require(SUCCEEDED(hr), std::format(L"Theme source '{}' should parse: {}", text, message)) && ok;
        if (SUCCEEDED(hr)) destination.emplace(std::move(key), std::move(source));
    };

    addSource(theme.palette, L"base", L"#336699");
    addSource(theme.palette, L"white", L"#FFFFFF");
    addSource(theme.palette, L"black", L"#000000");
    addSource(theme.palette, L"red", L"#FF0000");
    addSource(theme.palette, L"green", L"#00FF00");
    addSource(theme.palette, L"blue", L"#0000FF");
    addSource(theme.palette, L"lighter", L"lighten(palette.base,20%)");
    addSource(theme.palette, L"darker", L"darken(palette.base,20%)");
    addSource(theme.palette, L"mixed", L"blend(palette.black,palette.white,25%)");
    addSource(theme.palette, L"readable", L"contrast(palette.base,palette.white,palette.black)");
    addSource(theme.palette, L"tone60", L"perceptualTone(palette.base,60)");
    addSource(theme.palette, L"contrastFixed", L"ensureContrast(palette.base,palette.white,4.5)");
    addSource(theme.palette, L"harmonized", L"harmonize(palette.base,palette.red,25%)");

    addSource(theme.colors, L"app.accent", L"systemAccent()");
    addSource(theme.colors, L"window.background", L"systemColor(window)");
    addSource(theme.colors, L"menu.background", L"tone(palette.white,palette.black)");
    addSource(theme.colors, L"menu.text", L"ref(palette.readable)");
    addSource(theme.colors, L"menu.disabledText", L"alpha(palette.base,50%)");
    addSource(theme.colors, L"folderView.itemBackgroundSelected", L"seededChoice(runtime.seed,palette.red,palette.green,palette.blue)");

    auto context = Common::Settings::MakeSystemThemeResolutionContext(true);
    Common::Settings::ResolvedThemeColors resolved;
    std::wstring resolveMessage;
    ok = Require(SUCCEEDED(Common::Settings::ResolveThemeDefinition(theme, context, resolved, &resolveMessage)),
                 std::format(L"Complete version 2 expression graph should resolve: {}", resolveMessage)) && ok;
    ok = Require(resolved.colors[L"menu.background"] == 0xFF000000u, L"tone() should select the dark reference for a dark base.") && ok;
    ok = Require(resolved.colors[L"menu.disabledText"] == 0x80336699u, L"alpha() should replace alpha deterministically.") && ok;
    ok = Require(resolved.dynamicColors.size() == 1u, L"An allowlisted paint-time source should compile without being parsed during paint.") && ok;

    if (const auto choice = resolved.dynamicColors.find(L"folderView.itemBackgroundSelected"); choice != resolved.dynamicColors.end())
    {
        const uint32_t selected = Common::Settings::EvaluateDynamicThemeColor(
            choice->second, Common::Settings::ThemeRuntimeContext{.seedHash32 = 1u, .highContrast = false});
        ok = Require(selected == 0xFF00FF00u, L"seededChoice() should use stable seed modulo candidate order.") && ok;
        const uint32_t suppressed = Common::Settings::EvaluateDynamicThemeColor(
            choice->second, Common::Settings::ThemeRuntimeContext{.seedHash32 = 1u, .highContrast = true});
        ok = Require(suppressed == choice->second.fallbackArgb, L"High Contrast should suppress paint-time theme sources.") && ok;
    }

    Common::Settings::ThemeDefinition rainbowTheme = theme;
    rainbowTheme.colors.erase(L"folderView.itemBackgroundSelected");
    addSource(rainbowTheme.colors, L"folderView.itemBackgroundSelected", L"seededRainbow(runtime.seed,70%,90%,100%,15)");
    Common::Settings::ResolvedThemeColors rainbowResolved;
    ok = Require(SUCCEEDED(Common::Settings::ResolveThemeDefinition(rainbowTheme, context, rainbowResolved, &resolveMessage)),
                 std::format(L"seededRainbow() should resolve on the allowlisted selection token: {}", resolveMessage)) && ok;
    if (const auto rainbow = rainbowResolved.dynamicColors.find(L"folderView.itemBackgroundSelected"); rainbow != rainbowResolved.dynamicColors.end())
    {
        const Common::Settings::ThemeRuntimeContext runtime{.seedHash32 = 42u, .highContrast = false};
        const uint32_t first = Common::Settings::EvaluateDynamicThemeColor(rainbow->second, runtime);
        const uint32_t second = Common::Settings::EvaluateDynamicThemeColor(rainbow->second, runtime);
        ok = Require(first == second && first != rainbow->second.fallbackArgb, L"seededRainbow() should be stable for the same seed.") && ok;
    }

    Common::Settings::ThemeDefinition rejectedDynamic = theme;
    addSource(rejectedDynamic.colors, L"navigation.accent", L"seededChoice(runtime.seed,palette.red,palette.green)");
    Common::Settings::ResolvedThemeColors rejectedResolved;
    ok = Require(FAILED(Common::Settings::ResolveThemeDefinition(rejectedDynamic, context, rejectedResolved, &resolveMessage)) &&
                     resolveMessage.find(L"does not accept") != std::wstring::npos,
                 L"Paint-time sources should be rejected outside the explicit semantic-token allowlist.") && ok;

    Common::Settings::ThemeDefinition referencedDynamic = theme;
    referencedDynamic.colors.erase(L"app.accent");
    addSource(referencedDynamic.colors, L"app.accent", L"ref(folderView.itemBackgroundSelected)");
    Common::Settings::ResolvedThemeColors referencedDynamicResolved;
    ok = Require(FAILED(Common::Settings::ResolveThemeDefinition(referencedDynamic, context, referencedDynamicResolved, &resolveMessage)) &&
                     resolveMessage.find(L"cannot be referenced") != std::wstring::npos,
                 L"Static sources should reject forward references to paint-time programs.") && ok;

    Common::Settings::ThemeColorSource invalidRuntime;
    ok = Require(FAILED(Common::Settings::ParseThemeColorSource(L"seededChoice(runtime.seed,palette.red)", invalidRuntime)),
                 L"seededChoice() should require at least two candidates.") && ok;
    ok = Require(FAILED(Common::Settings::ParseThemeColorSource(
                     L"seededChoice(runtime.seed,palette.red,palette.green,palette.blue,palette.white,palette.black,palette.base,palette.lighter,palette.darker,palette.mixed)",
                     invalidRuntime)),
                 L"seededChoice() should reject more than eight candidates.") && ok;

    Common::Settings::ThemeColorSource systemAccent;
    ok = Require(SUCCEEDED(Common::Settings::ParseThemeColorSource(L"systemAccent()", systemAccent)) &&
                     Common::Settings::FormatThemeColorSource(systemAccent) == L"systemAccent()",
                 L"systemAccent() should preserve its compatibility spelling on round trip.") && ok;

    constexpr std::array canonicalExpressions{
        std::pair{std::wstring_view(L"ref(palette.base)"), std::wstring_view(L"ref(palette.base)")},
        std::pair{std::wstring_view(L"lighten(palette.base,20%)"), std::wstring_view(L"lighten(palette.base,0.2)")},
        std::pair{std::wstring_view(L"darken(palette.base,20%)"), std::wstring_view(L"darken(palette.base,0.2)")},
        std::pair{std::wstring_view(L"alpha(palette.base,50%)"), std::wstring_view(L"alpha(palette.base,0.5)")},
        std::pair{std::wstring_view(L"blend(palette.black,palette.white,25%)"), std::wstring_view(L"blend(palette.black,palette.white,0.25)")},
        std::pair{std::wstring_view(L"contrast(palette.base)"), std::wstring_view(L"contrast(palette.base)")},
        std::pair{std::wstring_view(L"contrast(palette.base,palette.white,palette.black)"),
                  std::wstring_view(L"contrast(palette.base,palette.white,palette.black)")},
        std::pair{std::wstring_view(L"perceptualTone(palette.base,60)"), std::wstring_view(L"perceptualTone(palette.base,60)")},
        std::pair{std::wstring_view(L"ensureContrast(palette.base,palette.white,4.5)"),
                  std::wstring_view(L"ensureContrast(palette.base,palette.white,4.5)")},
        std::pair{std::wstring_view(L"harmonize(palette.base,palette.red,25%)"),
                  std::wstring_view(L"harmonize(palette.base,palette.red,0.25)")},
        std::pair{std::wstring_view(L"systemAccent()"), std::wstring_view(L"systemAccent()")},
        std::pair{std::wstring_view(L"systemColor(window)"), std::wstring_view(L"systemColor(window)")},
        std::pair{std::wstring_view(L"tone(palette.white,palette.black)"), std::wstring_view(L"tone(palette.white,palette.black)")},
        std::pair{std::wstring_view(L"seededRainbow(runtime.seed,70%,90%,100%,15)"),
                  std::wstring_view(L"seededRainbow(runtime.seed,0.7,0.9,1,15)")},
        std::pair{std::wstring_view(L"seededChoice(runtime.seed,palette.red,palette.green)"),
                  std::wstring_view(L"seededChoice(runtime.seed,palette.red,palette.green)")},
    };
    for (const auto& [input, expected] : canonicalExpressions)
    {
        Common::Settings::ThemeColorSource source;
        const HRESULT parseHr = Common::Settings::ParseThemeColorSource(input, source);
        ok = Require(SUCCEEDED(parseHr) && Common::Settings::FormatThemeColorSource(source) == expected,
                     std::format(L"Theme expression '{}' should format as canonical schema value '{}'.", input, expected)) && ok;
    }

    Common::Settings::ThemeDefinition alphaAwareTheme;
    alphaAwareTheme.id          = L"user/alpha-aware-contrast";
    alphaAwareTheme.name        = L"Alpha Aware Contrast";
    alphaAwareTheme.baseThemeId = L"builtin/dark";
    addSource(alphaAwareTheme.palette, L"halfBlack", L"#80000000");
    addSource(alphaAwareTheme.palette, L"white", L"#FFFFFFFF");
    addSource(alphaAwareTheme.colors, L"app.accent", L"ensureContrast(palette.halfBlack,palette.white,7)");
    Common::Settings::ResolvedThemeColors alphaAwareResolved;
    ok = Require(FAILED(Common::Settings::ResolveThemeDefinition(alphaAwareTheme, context, alphaAwareResolved, &resolveMessage)),
                 L"ensureContrast() should reject a translucent foreground that cannot meet the requested rendered ratio.") && ok;

    alphaAwareTheme.palette.clear();
    alphaAwareTheme.colors.clear();
    addSource(alphaAwareTheme.palette, L"halfWhite", L"#80FFFFFF");
    addSource(alphaAwareTheme.palette, L"black", L"#FF000000");
    addSource(alphaAwareTheme.colors, L"app.accent", L"ensureContrast(palette.halfWhite,palette.black,4.5)");
    ok = Require(SUCCEEDED(Common::Settings::ResolveThemeDefinition(alphaAwareTheme, context, alphaAwareResolved, &resolveMessage)) &&
                     (alphaAwareResolved.colors[L"app.accent"] & 0xFF000000u) == 0x80000000u,
                 L"ensureContrast() should measure the rendered composite and preserve alpha when the requested ratio is attainable.") && ok;

    alphaAwareTheme.palette.clear();
    alphaAwareTheme.colors.clear();
    addSource(alphaAwareTheme.palette, L"black", L"#FF000000");
    addSource(alphaAwareTheme.palette, L"halfWhite", L"#80FFFFFF");
    addSource(alphaAwareTheme.colors, L"app.accent", L"ensureContrast(palette.black,palette.halfWhite,4.5)");
    ok = Require(FAILED(Common::Settings::ResolveThemeDefinition(alphaAwareTheme, context, alphaAwareResolved, &resolveMessage)),
                 L"ensureContrast() should reject a translucent background because its rendered backdrop is unknown.") && ok;

    Common::Settings::ThemeDefinition cyclic;
    cyclic.id = L"user/cyclic";
    cyclic.name = L"Cyclic";
    cyclic.baseThemeId = L"builtin/dark";
    addSource(cyclic.colors, L"app.accent", L"ref(menu.background)");
    addSource(cyclic.colors, L"menu.background", L"ref(app.accent)");
    Common::Settings::ResolvedThemeColors cycleResult;
    ok = Require(FAILED(Common::Settings::ResolveThemeDefinition(cyclic, context, cycleResult, &resolveMessage)) &&
                     resolveMessage.find(L"cycle") != std::wstring::npos,
                 L"Theme dependency cycles should fail with a useful diagnostic.") && ok;

    return ok;
}

[[nodiscard]] bool TestThemeExpressionNumbersAreLocaleInvariant()
{
    const wchar_t* previousLocaleText = _wsetlocale(LC_NUMERIC, nullptr);
    const std::wstring previousLocale = previousLocaleText ? previousLocaleText : L"C";
    const int previousThreadLocaleMode = _configthreadlocale(_ENABLE_PER_THREAD_LOCALE);
    const auto restoreLocale = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(_wsetlocale(LC_NUMERIC, previousLocale.c_str()));
        static_cast<void>(_configthreadlocale(previousThreadLocaleMode));
    });

    const auto parseEnsureContrast = [](std::wstring_view number, double& parsed) noexcept
    {
        Common::Settings::ThemeColorSource source;
        const std::wstring expression = std::format(L"ensureContrast(app.accent,menu.text,{})", number);
        if (FAILED(Common::Settings::ParseThemeColorSource(expression, source)))
        {
            return false;
        }
        parsed = source.parameters[0];
        return true;
    };

    bool ok = true;
    ok = Require(previousThreadLocaleMode != -1, L"RedConfigureTests should be able to isolate LC_NUMERIC to the current thread.") && ok;
    ok = Require(_wsetlocale(LC_NUMERIC, L"C") != nullptr, L"RedConfigureTests should activate the C numeric locale.") && ok;

    double cLocaleValue = 0.0;
    ok = Require(parseEnsureContrast(L"4.5", cLocaleValue) && cLocaleValue == 4.5,
                 L"Theme expression should parse a dot-decimal number in the C locale.") &&
         ok;

    const std::array commaLocaleCandidates{L"French_France.1252", L"French_France", L"fr-FR"};
    const wchar_t* commaLocale = nullptr;
    for (const wchar_t* candidate : commaLocaleCandidates)
    {
        commaLocale = _wsetlocale(LC_NUMERIC, candidate);
        if (commaLocale != nullptr)
        {
            break;
        }
    }
    ok = Require(commaLocale != nullptr, L"RedConfigureTests could not activate an installed comma-decimal locale.") && ok;

    if (commaLocale != nullptr)
    {
        double commaLocaleValue = 0.0;
        ok = Require(parseEnsureContrast(L"4.5", commaLocaleValue) && commaLocaleValue == cLocaleValue,
                     L"Theme dot-decimal parsing should be identical under a comma-decimal locale.") &&
             ok;
        ok = Require(! parseEnsureContrast(L"4,5", commaLocaleValue),
                     L"Theme expressions should reject comma decimals under every process locale.") &&
             ok;
    }

    double ignored = 0.0;
    ok = Require(parseEnsureContrast(L"+4.5", ignored) && ignored == 4.5, L"Theme numeric grammar should accept an explicit leading plus sign.") && ok;
    ok = Require(! parseEnsureContrast(L"nan", ignored), L"Theme numeric grammar should reject NaN.") && ok;
    ok = Require(! parseEnsureContrast(L"inf", ignored), L"Theme numeric grammar should reject infinity.") && ok;
    ok = Require(! parseEnsureContrast(L"1e309", ignored), L"Theme numeric grammar should reject overflow.") && ok;
    ok = Require(! parseEnsureContrast(std::wstring(65u, L'1'), ignored), L"Theme numeric grammar should enforce its bounded token length.") && ok;
    return ok;
}

[[nodiscard]] bool TestSettingsStoreThemeDirectoryUsesSharedParser()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureSettingsThemeParserParityTest", ec);
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for SettingsStore theme parser parity tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    bool ok = true;
    ok      = Require(WriteTestTextFile(tempRoot / L"Good.theme.json5",
                                        R"json5({
  formatVersion: 2,
  id: "user/settings-good",
  name: "Settings Good",
  baseThemeId: "builtin/dark",
  colors: {
    "app.accent": "#336699",
  },
})json5"),
                      L"Failed to write valid SettingsStore theme fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"InvalidId.theme.json5",
                                        R"json5({
  formatVersion: 2,
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

[[nodiscard]] bool TestShippedThemesUseVersion2Sources()
{
    constexpr std::array<std::pair<std::wstring_view, uint64_t>, 6> kMigratedThemeGoldens{{
        {L"user/forest-mist", 0xB34ACE65892AE638ull},
        {L"user/neon-tokyo", 0xE1B5EF92856E3518ull},
        {L"user/paper-and-ink", 0x2F4D788E11B441BFull},
        {L"user/retro-terminal", 0xADD3A34B68A03402ull},
        {L"user/solar-flare", 0x42B6B9537475EBD2ull},
        {L"user/ugly", 0xEC2CC5EBB3A3C727ull},
    }};
    const auto linear = [](uint8_t channel) noexcept
    {
        const double value = static_cast<double>(channel) / 255.0;
        return value <= 0.04045 ? value / 12.92 : std::pow((value + 0.055) / 1.055, 2.4);
    };
    const auto luminance = [&](uint32_t argb) noexcept
    {
        return 0.2126 * linear(static_cast<uint8_t>((argb >> 16u) & 0xFFu)) +
               0.7152 * linear(static_cast<uint8_t>((argb >> 8u) & 0xFFu)) +
               0.0722 * linear(static_cast<uint8_t>(argb & 0xFFu));
    };
    const auto contrastRatio = [&](uint32_t first, uint32_t second) noexcept
    {
        const double firstLuminance = luminance(first);
        const double secondLuminance = luminance(second);
        return (std::max(firstLuminance, secondLuminance) + 0.05) / (std::min(firstLuminance, secondLuminance) + 0.05);
    };

    std::error_code ec;
    std::filesystem::path root = std::filesystem::current_path(ec);
    if (ec)
    {
        return Require(false, L"Could not resolve the current path for shipped theme validation.");
    }

    std::filesystem::path themesDirectory;
    for (size_t depth = 0u; depth < 8u && ! root.empty(); ++depth)
    {
        const std::filesystem::path candidate = root / L"Specs" / L"Themes";
        if (std::filesystem::is_directory(candidate, ec) && ! ec)
        {
            themesDirectory = candidate;
            break;
        }
        ec.clear();
        root = root.parent_path();
    }
    bool ok = Require(! themesDirectory.empty(), L"Could not locate Specs\\Themes for shipped theme validation.");
    if (themesDirectory.empty()) return false;

    size_t sourceFileCount = 0u;
    for (std::filesystem::directory_iterator it(themesDirectory, ec), end; ! ec && it != end; it.increment(ec))
    {
        if (it->is_regular_file(ec) && ! ec && it->path().filename().wstring().ends_with(L".theme.json5")) ++sourceFileCount;
        ec.clear();
    }
    ok = Require(! ec, L"Could not enumerate shipped theme source files.") && ok;

    std::vector<Common::Settings::ThemeDefinition> themes;
    ok = Require(SUCCEEDED(Common::Settings::LoadThemeDefinitionsFromDirectory(themesDirectory, themes)),
                 L"The shipped theme directory should load through the shared strict parser.") && ok;
    ok = Require(themes.size() == sourceFileCount, L"Every shipped .theme.json5 file should be a valid version 2 theme.") && ok;
    ok = Require(themes.size() == 10u, L"The shipped catalogue should contain the six migrated themes, Dracula, and three Catppuccin themes.") && ok;

    bool foundDracula = false;
    bool foundLatte = false;
    bool foundFrappe = false;
    bool foundMocha = false;
    for (const Common::Settings::ThemeDefinition& theme : themes)
    {
        ok = Require(theme.formatVersion == 2u, std::format(L"Shipped theme '{}' should require formatVersion 2.", theme.id)) && ok;
        Common::Settings::ResolvedThemeColors resolved;
        std::wstring message;
        const bool effectiveDark = theme.baseThemeId != L"builtin/light";
        auto context = Common::Settings::MakeSystemThemeResolutionContext(effectiveDark);
        ok = Require(SUCCEEDED(Common::Settings::ResolveThemeDefinition(theme, context, resolved, &message)),
                     std::format(L"Shipped theme '{}' should resolve: {}", theme.id, message)) && ok;
        constexpr std::array<std::pair<std::wstring_view, std::wstring_view>, 4> kRequiredTextPairs{{
            {L"menu.text", L"menu.background"},
            {L"navigation.text", L"navigation.background"},
            {L"folderView.textNormal", L"folderView.background"},
            {L"monitor.textView.fg", L"monitor.textView.bg"},
        }};
        for (const auto& [foregroundKey, backgroundKey] : kRequiredTextPairs)
        {
            const auto foreground = resolved.colors.find(std::wstring(foregroundKey));
            const auto background = resolved.colors.find(std::wstring(backgroundKey));
            ok = Require(foreground != resolved.colors.end() && background != resolved.colors.end() &&
                             contrastRatio(foreground->second, background->second) >= 4.5,
                         std::format(L"Shipped theme '{}' should keep {} readable over {} at WCAG 4.5:1.", theme.id, foregroundKey, backgroundKey)) && ok;
        }
        const auto golden = std::ranges::find_if(kMigratedThemeGoldens, [&](const auto& entry) noexcept { return entry.first == theme.id; });
        if (golden != kMigratedThemeGoldens.end())
        {
            std::vector<std::pair<std::wstring, uint32_t>> colors(resolved.colors.begin(), resolved.colors.end());
            std::ranges::sort(colors, {}, &std::pair<std::wstring, uint32_t>::first);
            uint64_t hash = 14695981039346656037ull;
            const auto appendByte = [&](uint8_t value) noexcept
            {
                hash ^= value;
                hash *= 1099511628211ull;
            };
            for (const auto& [key, argb] : colors)
            {
                for (const wchar_t ch : key) appendByte(static_cast<uint8_t>(ch));
                appendByte(static_cast<uint8_t>('='));
                for (const char ch : std::format("{:08X}", argb)) appendByte(static_cast<uint8_t>(ch));
                appendByte(static_cast<uint8_t>('\n'));
            }
            ok = Require(colors.size() == 64u && hash == golden->second,
                         std::format(L"Migrated theme '{}' should remain effective-color-equivalent to its 64 pre-cutover values.", theme.id)) && ok;
        }
        foundDracula = foundDracula || theme.id == L"user/dracula";
        foundLatte = foundLatte || theme.id == L"user/catppuccin-latte";
        foundFrappe = foundFrappe || theme.id == L"user/catppuccin-frappe";
        foundMocha = foundMocha || theme.id == L"user/catppuccin-mocha";
    }
    ok = Require(foundDracula && foundLatte && foundFrappe && foundMocha,
                 L"Dracula and the Latte, Frappe, and Mocha Catppuccin themes should all ship.") && ok;
    return ok;
}

[[nodiscard]] bool TestSettingsStoreInlineThemesUseSharedParser()
{
    const auto uniqueTicks = std::chrono::steady_clock::now().time_since_epoch().count();
    const std::wstring appId = L"RedConfigureTests_InlineThemeParser_" + std::to_wstring(uniqueTicks);
    RemoveSettingsFiles(appId);
    const auto cleanup = wil::scope_exit([&] { RemoveSettingsFiles(appId); });

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(appId);
    bool ok = Require(! settingsPath.empty(), L"Settings path should be available for inline theme parser test.");
    ok      = Require(WriteTestTextFile(settingsPath,
                                        R"json({
  "schemaVersion": 16,
  "theme": {
    "themes": [
      {
        "formatVersion": 2,
        "id": "user/settings-good",
        "name": "Settings Good",
        "baseThemeId": "builtin/dark",
        "colors": {
          "app.accent": "#336699"
        }
      },
      {
        "formatVersion": 2,
        "id": "user/settings-malformed-key",
        "name": "Settings theme name that is deliberately longer than sixty-four UTF-16 code units for clamping",
        "baseThemeId": "future/vendor-theme",
        "colors": {
          "app.accent": "#445566",
          "not a theme key": "#112233",
          "app.background": "not-a-color"
        }
      },
      "keep-this-opaque",
      {
        "formatVersion": 2,
        "id": "user/settings-good",
        "name": "Settings Duplicate",
        "baseThemeId": "builtin/dark",
        "colors": {
          "app.accent": "#778899"
        }
      }
    ]
  }
})json"),
                      L"Failed to write inline settings theme fixture.") &&
              ok;

    Common::Settings::Settings loaded;
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(appId, loaded);
    ok                   = Require(SUCCEEDED(loadHr), L"Inline settings theme fixture should load without recovery.") && ok;
    ok                   = Require(loaded.theme.themes.size() == 2u,
                                   L"Inline settings themes should retain a theme with invalid colors while rejecting duplicate ids.") &&
         ok;
    const auto good = std::find_if(loaded.theme.themes.begin(), loaded.theme.themes.end(), [](const Common::Settings::ThemeDefinition& theme) noexcept
    { return theme.id == L"user/settings-good"; });
    if (good != loaded.theme.themes.end())
    {
        ok = Require(good->colors.size() == 1u && good->colors.contains(L"app.accent"), L"Inline settings theme should preserve valid color keys.") && ok;
    }
    else
    {
        ok = Require(false, L"Inline settings themes should preserve the first valid id.") && ok;
    }

    const auto lenient = std::find_if(loaded.theme.themes.begin(), loaded.theme.themes.end(), [](const Common::Settings::ThemeDefinition& theme) noexcept
    { return theme.id == L"user/settings-malformed-key"; });
    ok = Require(lenient != loaded.theme.themes.end(), L"One invalid color must not drop an otherwise usable inline theme.") && ok;
    if (lenient != loaded.theme.themes.end())
    {
        ok = Require(lenient->name.size() == 64u, L"Lenient inline theme parsing should clamp over-long names.") && ok;
        ok = Require(lenient->baseThemeId == L"future/vendor-theme", L"Lenient inline theme parsing should preserve unknown base theme ids.") && ok;
        ok = Require(lenient->colors.size() == 1u && lenient->colors.contains(L"app.accent"),
                     L"Lenient inline theme parsing should skip invalid color entries and preserve valid entries.") &&
             ok;
    }

    ok = Require(loaded.theme.opaqueThemeEntries.size() == 2u,
                 L"Structurally unusable and duplicate inline theme entries should be retained opaquely.") && ok;
    if (loaded.theme.opaqueThemeEntries.size() == 2u)
    {
        ok = Require(loaded.theme.opaqueThemeEntries[0].originalIndex == 2u && loaded.theme.opaqueThemeEntries[1].originalIndex == 3u,
                     L"Opaque inline theme entries should retain their original array positions.") && ok;
    }
    const HRESULT saveHr = Common::Settings::SaveSettings(appId, loaded);
    ok                   = Require(SUCCEEDED(saveHr), L"Failed to save lenient inline theme round-trip fixture.") && ok;
    Common::Settings::Settings reloaded;
    const HRESULT reloadHr = Common::Settings::TryLoadSettingsNoRecovery(appId, reloaded);
    ok = Require(SUCCEEDED(reloadHr), L"Failed to reload lenient inline theme round-trip fixture.") && ok;
    ok = Require(reloaded.theme.opaqueThemeEntries.size() == 2u,
                 L"Opaque and duplicate inline theme entries should survive save and reload.") && ok;
    if (reloaded.theme.opaqueThemeEntries.size() == 2u)
    {
        const auto& unusableEntry = reloaded.theme.opaqueThemeEntries[0];
        const auto* text          = std::get_if<std::string>(&unusableEntry.value.value);
        ok = Require(unusableEntry.originalIndex == 2u && text && *text == "keep-this-opaque",
                     L"Structurally unusable inline theme entry changed position or value during save and reload.") && ok;

        const auto& duplicateEntry = reloaded.theme.opaqueThemeEntries[1];
        const auto* object = std::get_if<Common::Settings::JsonValue::ObjectPtr>(&duplicateEntry.value.value);
        bool duplicateNamePreserved = false;
        if (object && *object)
        {
            for (const auto& [key, value] : (*object)->members)
            {
                const auto* stringValue = std::get_if<std::string>(&value.value);
                duplicateNamePreserved = duplicateNamePreserved || (key == "name" && stringValue && *stringValue == "Settings Duplicate");
            }
        }
        ok = Require(duplicateEntry.originalIndex == 3u && duplicateNamePreserved,
                     L"Duplicate inline theme entry changed position or authored content during save and reload.") && ok;
    }

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
    ok      = Require(ValidatePlaceholders(L"Value {0}", L"Valeur {}").status == RedConfigure::Localization::PlaceholderStatus::BarePlaceholder,
                      L"Bare placeholders should be rejected.") &&
              ok;
    ok      = Require(ValidatePlaceholders(L"Value {0:08X}", L"Valeur {:08X}").status == RedConfigure::Localization::PlaceholderStatus::UnindexedFormatSpec,
                      L"Unindexed format specs should be rejected.") &&
              ok;
    ok      = Require(ValidatePlaceholders(L"Value {0}", L"Valeur").status == RedConfigure::Localization::PlaceholderStatus::PlaceholderMismatch,
                      L"Missing target placeholders should be rejected.") &&
              ok;
    ok      = Require(ValidatePlaceholders(L"Values {0} {0}", L"Valeurs {0}").status == RedConfigure::Localization::PlaceholderStatus::PlaceholderMismatch,
                      L"Duplicated source placeholders must not collapse to one target placeholder.") &&
              ok;
    ok      = Require(ValidatePlaceholders(L"Value {0}", L"Valeur {0} {0}").status == RedConfigure::Localization::PlaceholderStatus::PlaceholderMismatch,
                      L"Duplicated target placeholders must not collapse to one source placeholder.") &&
              ok;
    ok      = Require(ValidatePlaceholders(L"Value {0}", L"Valeur %s").status == RedConfigure::Localization::PlaceholderStatus::PrintfPlaceholder,
                      L"Printf placeholders should be rejected.") &&
              ok;
    ok = Require(ValidatePlaceholders(L"Value {0}", L"Valeur {0").status == RedConfigure::Localization::PlaceholderStatus::InvalidPlaceholder,
                 L"Expected an unbalanced opening brace to be rejected.") && ok;
    ok = Require(ValidatePlaceholders(L"Value {0}", L"Valeur {name}").status == RedConfigure::Localization::PlaceholderStatus::InvalidPlaceholder,
                 L"Expected a non-positional named placeholder to be rejected.") && ok;
    return ok;
}

[[nodiscard]] bool TestTranslationViewSearchFilterAndSort()
{
    using RedConfigure::Localization::PlaceholderStatus;
    bool ok = true;

    std::vector<RedConfigure::TranslationEntry> rows;
    rows.push_back(RedConfigure::TranslationEntry{.id         = L"IDS_SAVE",
                                                  .sourceText = L"Save file",
                                                  .targetText = L"Enregistrer",
                                                  .validation = RedConfigure::Localization::PlaceholderValidationResult{.status = PlaceholderStatus::Ok}});
    rows.push_back(RedConfigure::TranslationEntry{.id         = L"IDS_OPEN",
                                                  .sourceText = L"Open file",
                                                  .targetText = L"Ouvrir",
                                                  .validation = RedConfigure::Localization::PlaceholderValidationResult{.status = PlaceholderStatus::Ok}});
    rows.push_back(RedConfigure::TranslationEntry{
        .id         = L"IDS_DELETE",
        .sourceText = L"Delete {0}",
        .targetText = L"Supprimer",
        .validation = RedConfigure::Localization::PlaceholderValidationResult{.status = PlaceholderStatus::PlaceholderMismatch}});

    RedConfigure::LocalizationViewOptions options;
    options.searchText       = L"file";
    std::vector<size_t> view = RedConfigure::BuildTranslationView(rows, options);
    ok                       = Require(view == std::vector<size_t>{0u, 1u}, L"Translation view search should match source and preserve source order.") && ok;

    options.searchText   = {};
    options.idFilterText = L"delete";
    options.statusFilter = RedConfigure::LocalizationStatusFilter::Problems;
    view                 = RedConfigure::BuildTranslationView(rows, options);
    ok                   = Require(view == std::vector<size_t>{2u}, L"Translation view filters should combine ID and status filters.") && ok;

    options.idFilterText  = {};
    options.statusFilter  = RedConfigure::LocalizationStatusFilter::All;
    options.sortColumn    = RedConfigure::LocalizationViewColumn::Target;
    options.sortDirection = RedConfigure::LocalizationSortDirection::Ascending;
    view                  = RedConfigure::BuildTranslationView(rows, options);
    ok                    = Require(view == std::vector<size_t>{0u, 1u, 2u}, L"Translation view target sort should order by target text ascending.") && ok;

    options.sortDirection = RedConfigure::LocalizationSortDirection::Descending;
    view                  = RedConfigure::BuildTranslationView(rows, options);
    ok                    = Require(view == std::vector<size_t>{2u, 1u, 0u}, L"Translation view target sort should reverse when descending.") && ok;

    return ok;
}

[[nodiscard]] bool ContainsCulture(std::span<const std::wstring> cultures, std::wstring_view culture) noexcept
{
    for (const std::wstring& candidate : cultures)
    {
        if (candidate == culture)
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] const RedConfigure::LocalizationReviewRow* FindReviewRow(std::span<const RedConfigure::LocalizationReviewRow> rows,
                                                                       std::wstring_view ownerName,
                                                                       std::wstring_view id) noexcept
{
    for (const RedConfigure::LocalizationReviewRow& row : rows)
    {
        if (row.ownerName == ownerName && row.id == id)
        {
            return &row;
        }
    }
    return nullptr;
}

[[nodiscard]] const RedConfigure::LocalizationTargetCell* FindReviewTargetCell(const RedConfigure::LocalizationReviewRow& row,
                                                                               std::wstring_view cultureName) noexcept
{
    for (const RedConfigure::LocalizationTargetCell& cell : row.targets)
    {
        if (cell.cultureName == cultureName)
        {
            return &cell;
        }
    }
    return nullptr;
}

[[nodiscard]] bool WriteLocalizationReviewWorkspaceFixture(const std::filesystem::path& tempRoot)
{
    bool ok = true;
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"App.vcxproj",
                                        R"xml(<?xml version="1.0" encoding="utf-8"?>
<Project>
  <ItemGroup>
    <ResourceCompile Include="App.rc" />
  </ItemGroup>
</Project>)xml"),
                      L"Failed to write app localization review project fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"App.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Hello {0}"
    IDS_APP_ONLY "Only app"
END
)rc"),
                      L"Failed to write app localization review source fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"Lang" / L"fr-FR" / L"App-fr-FR.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Bonjour {0}"
END
)rc"),
                      L"Failed to write app fr-FR localization review fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"Lang" / L"cs-CZ" / L"App-cs-CZ.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Ahoj {0}"
END
)rc"),
                      L"Failed to write app cs-CZ localization review fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Plugin" / L"Plugin.vcxproj",
                                        R"xml(<?xml version="1.0" encoding="utf-8"?>
<Project>
  <ItemGroup>
    <ResourceCompile Include="Plugin.rc" />
  </ItemGroup>
</Project>)xml"),
                      L"Failed to write plugin localization review project fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Plugin" / L"Plugin.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Plugin hello {0}"
    IDS_PLUGIN_ONLY "Only plugin"
END
)rc"),
                      L"Failed to write plugin localization review source fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Plugin" / L"Lang" / L"fr-FR" / L"Plugin-fr-FR.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Plugin bonjour"
    IDS_PLUGIN_ONLY "Plugin seulement"
END
)rc"),
                      L"Failed to write plugin fr-FR localization review fixture.") &&
              ok;
    return ok;
}

[[nodiscard]] bool TestRedConfigureSessionBuildsAllOwnerAllCultureLocalizationReview()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureLocalizationReviewTest", ec);
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for localization review tests.");
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
                      L"Failed to write app localization review project fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"App.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Hello {0}"
    IDS_APP_ONLY "Only app"
END
)rc"),
                      L"Failed to write app localization review source fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"Lang" / L"fr-FR" / L"App-fr-FR.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Bonjour {0}"
END
)rc"),
                      L"Failed to write app fr-FR localization review fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"Lang" / L"cs-CZ" / L"App-cs-CZ.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Ahoj {0}"
END
)rc"),
                      L"Failed to write app cs-CZ localization review fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Plugin" / L"Plugin.vcxproj",
                                        R"xml(<?xml version="1.0" encoding="utf-8"?>
<Project>
  <ItemGroup>
    <ResourceCompile Include="Plugin.rc" />
  </ItemGroup>
</Project>)xml"),
                      L"Failed to write plugin localization review project fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Plugin" / L"Plugin.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Plugin hello {0}"
    IDS_PLUGIN_ONLY "Only plugin"
END
)rc"),
                      L"Failed to write plugin localization review source fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Plugin" / L"Lang" / L"fr-FR" / L"Plugin-fr-FR.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_PLUGIN_ONLY "Plugin seulement"
END
)rc"),
                      L"Failed to write plugin fr-FR localization review fixture.") &&
              ok;

    RedConfigure::RedConfigureSession session;
    ok = Require(SUCCEEDED(session.LoadWorkspace(tempRoot, L"fr-FR")), L"Session should load the localization review fixture.") && ok;

    const auto reviewRows     = session.GetLocalizationReviewRows();
    const auto reviewCultures = session.GetLocalizationReviewCultures();
    ok                        = Require(reviewRows.size() == 4u, L"Review rows should include all source string IDs from all owners.") && ok;
    ok                        = Require(ContainsCulture(reviewCultures, L"fr-FR"), L"Review cultures should include fr-FR.") && ok;
    ok                        = Require(ContainsCulture(reviewCultures, L"cs-CZ"), L"Review cultures should include cs-CZ.") && ok;
    if (reviewRows.size() == 4u)
    {
        ok = Require(reviewRows[0u].ownerName == L"App" && reviewRows[0u].id == L"IDS_HELLO" && reviewRows[1u].ownerName == L"App" &&
                         reviewRows[1u].id == L"IDS_APP_ONLY" && reviewRows[2u].ownerName == L"Plugin" && reviewRows[2u].id == L"IDS_HELLO" &&
                         reviewRows[3u].ownerName == L"Plugin" && reviewRows[3u].id == L"IDS_PLUGIN_ONLY",
                     L"Parallel localization review loading should preserve deterministic owner/source row order.") &&
             ok;
    }

    const auto* appOnly = FindReviewRow(reviewRows, L"App", L"IDS_APP_ONLY");
    ok                  = Require(appOnly != nullptr, L"Review rows should include app-only string IDs.") && ok;
    if (appOnly)
    {
        ok = Require(appOnly->sourceText == L"Only app", L"Review rows should preserve English source text.") && ok;
    }

    const auto* appHello = FindReviewRow(reviewRows, L"App", L"IDS_HELLO");
    ok                   = Require(appHello != nullptr, L"Review rows should include app shared string IDs.") && ok;
    if (appHello)
    {
        const auto* frCell = FindReviewTargetCell(*appHello, L"fr-FR");
        const auto* csCell = FindReviewTargetCell(*appHello, L"cs-CZ");
        ok                 = Require(frCell != nullptr && frCell->targetText == L"Bonjour {0}", L"Review rows should merge existing fr-FR target text.") && ok;
        ok                 = Require(csCell != nullptr && csCell->targetText == L"Ahoj {0}", L"Review rows should merge existing cs-CZ target text.") && ok;
    }

    const auto* pluginOnly = FindReviewRow(reviewRows, L"Plugin", L"IDS_PLUGIN_ONLY");
    ok                     = Require(pluginOnly != nullptr, L"Review rows should include plugin-only string IDs.") && ok;
    if (pluginOnly)
    {
        const auto* frCell = FindReviewTargetCell(*pluginOnly, L"fr-FR");
        const auto* csCell = FindReviewTargetCell(*pluginOnly, L"cs-CZ");
        ok = Require(frCell != nullptr && frCell->targetText == L"Plugin seulement", L"Review rows should merge plugin fr-FR target text.") && ok;
        ok = Require(csCell != nullptr && csCell->targetText == L"Only plugin" && ! csCell->hasExistingTranslation,
                     L"Missing target cultures should display English source text without marking an existing translation.") &&
             ok;
    }

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestLocalizationReviewViewFiltersSearchesAndSorts()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureLocalizationReviewViewTest", ec);
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for localization review view tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    bool ok = WriteLocalizationReviewWorkspaceFixture(tempRoot);

    RedConfigure::RedConfigureSession session;
    ok = Require(SUCCEEDED(session.LoadWorkspace(tempRoot, L"fr-FR")), L"Session should load the localization review view fixture.") && ok;

    RedConfigure::LocalizationReviewViewOptions options;
    options.visibleOwnerNames   = {L"Plugin"};
    options.visibleCultureNames = {L"fr-FR", L"cs-CZ"};
    std::vector<size_t> view    = RedConfigure::BuildLocalizationReviewView(session.GetLocalizationReviewRows(), options);
    ok                          = Require(view.size() == 2u, L"Review view owner filtering should keep only checked owners.") && ok;
    for (const size_t rowIndex : view)
    {
        ok = Require(session.GetLocalizationReviewRows()[rowIndex].ownerName == L"Plugin", L"Review view should not include unchecked owners.") && ok;
    }

    options.visibleOwnerNames = {L"App", L"Plugin"};
    options.searchText        = L"cs-CZ";
    view                      = RedConfigure::BuildLocalizationReviewView(session.GetLocalizationReviewRows(), options);
    ok                        = Require(view.size() == 4u, L"Review view search should match visible language names.") && ok;

    options.searchText        = L"Ahoj";
    view                      = RedConfigure::BuildLocalizationReviewView(session.GetLocalizationReviewRows(), options);
    ok                        = Require(view.size() == 1u && session.GetLocalizationReviewRows()[view.front()].ownerName == L"App",
                                        L"Review view search should match visible target-language text.") &&
                                ok;

    options.visibleCultureNames = {L"fr-FR"};
    view                        = RedConfigure::BuildLocalizationReviewView(session.GetLocalizationReviewRows(), options);
    ok                          = Require(view.empty(), L"Review view search should ignore hidden target-language text.") && ok;

    options.searchText   = {};
    options.statusFilter = RedConfigure::LocalizationStatusFilter::Problems;
    view                 = RedConfigure::BuildLocalizationReviewView(session.GetLocalizationReviewRows(), options);
    ok                   = Require(view.size() == 2u && session.GetLocalizationReviewRows()[view.front()].id == L"IDS_APP_ONLY" &&
                                       session.GetLocalizationReviewRows()[view.front()].ownerName == L"App" &&
                                       session.GetLocalizationReviewRows()[view.back()].id == L"IDS_HELLO" &&
                                       session.GetLocalizationReviewRows()[view.back()].ownerName == L"Plugin",
                                   L"Review view problem filtering should include placeholder problems and missing visible translations.") &&
                           ok;

    options.statusFilter    = RedConfigure::LocalizationStatusFilter::All;
    options.sortColumn      = RedConfigure::LocalizationViewColumn::Target;
    options.sortCultureName = L"fr-FR";
    options.sortDirection   = RedConfigure::LocalizationSortDirection::Descending;
    view                    = RedConfigure::BuildLocalizationReviewView(session.GetLocalizationReviewRows(), options);
    ok                      = Require(view.size() == 4u && session.GetLocalizationReviewRows()[view.front()].ownerName == L"Plugin",
                                      L"Review view target-culture sorting should use the requested visible culture.") &&
                              ok;

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool ContainsExportPreview(std::span<const RedConfigure::LocalizationExportPreview> previews,
                                         std::wstring_view ownerName,
                                         std::wstring_view cultureName,
                                         std::wstring_view expectedText) noexcept
{
    for (const RedConfigure::LocalizationExportPreview& preview : previews)
    {
        if (preview.ownerName == ownerName && preview.cultureName == cultureName && preview.text.find(expectedText) != std::wstring::npos)
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] bool TestLocalizationReviewEditingAndExportPreviews()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureLocalizationReviewEditExportTest", ec);
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for localization review edit/export tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    bool ok = WriteLocalizationReviewWorkspaceFixture(tempRoot);

    RedConfigure::RedConfigureSession session;
    ok = Require(SUCCEEDED(session.LoadWorkspace(tempRoot, L"fr-FR")), L"Session should load the localization review edit/export fixture.") && ok;

    const auto* appHello = FindReviewRow(session.GetLocalizationReviewRows(), L"App", L"IDS_HELLO");
    ok                   = Require(appHello != nullptr, L"Edit/export fixture should expose App IDS_HELLO.") && ok;
    if (appHello)
    {
        const size_t rowIndex = static_cast<size_t>(appHello - session.GetLocalizationReviewRows().data());
        ok                    = Require(session.UpdateLocalizationReviewTarget(rowIndex, L"cs-CZ", L"Ahoj upraveno {0}"),
                                        L"Review target edits should accept valid placeholder-equivalent text.") &&
                                ok;
        ok                    = Require(! session.UpdateLocalizationReviewTarget(rowIndex, L"cs-CZ", L"Ahoj upraveno"),
                                        L"Review target edits should reject placeholder mismatches.") &&
                                ok;
        ok = Require(session.CanUndo() && session.Undo(), L"Localization edits should be undoable from the global command model.") && ok;
        ok = Require(session.CanRedo() && session.Redo(), L"Undone localization edits should be redoable from the global command model.") && ok;

        size_t csCultureIndex = 0u;
        const auto currentRows = session.GetLocalizationReviewRows();
        if (rowIndex < currentRows.size())
        {
            for (size_t index = 0u; index < currentRows[rowIndex].targets.size(); ++index)
            {
                if (currentRows[rowIndex].targets[index].cultureName == L"cs-CZ")
                {
                    csCultureIndex = index;
                    break;
                }
            }
        }
        ok = Require(session.ApplyClipboardMatrix(rowIndex, csCultureIndex, L"Clipboard {0}"),
                     L"Localization matrix should paste rectangular clipboard text into the selected culture.") &&
             ok;
        ok = Require(session.Undo(), L"Rectangular clipboard paste should be undoable as one operation.") && ok;

        const auto staleBatch = RedConfigure::Workflow::PreviewLocalizationBatch(
            session.GetLocalizationReviewRows(),
            {.kind = RedConfigure::Workflow::LocalizationBatchKind::FindReplace,
             .targetCulture = L"cs-CZ",
             .findText = L"upraveno",
             .replaceText = L"batch",
             .rowIndices = {rowIndex}});
        ok = Require(staleBatch.result == RedConfigure::Workflow::BatchApprovalResult::Ready,
                     L"Localization stale-approval fixture should produce a ready preview.") &&
             ok;
        ok = Require(session.UpdateLocalizationReviewTarget(rowIndex, L"cs-CZ", L"Novější {0}"),
                     L"Localization stale-approval fixture should accept the intervening edit.") &&
             ok;
        ok = Require(session.ApplyLocalizationBatch(staleBatch) == RedConfigure::Workflow::BatchApprovalResult::Stale,
                     L"Localization batch apply should reject an intervening target edit as stale.") &&
             ok;
        const auto* staleRow = FindReviewRow(session.GetLocalizationReviewRows(), L"App", L"IDS_HELLO");
        const auto* staleCell = staleRow ? FindReviewTargetCell(*staleRow, L"cs-CZ") : nullptr;
        ok = Require(staleCell != nullptr && staleCell->targetText == L"Novější {0}",
                     L"Stale localization rejection should preserve the newer target with zero mutation.") &&
             ok;
        ok = Require(session.Undo(), L"The intervening localization edit should remain the only Undo entry after stale rejection.") && ok;

        const auto readyBatch = RedConfigure::Workflow::PreviewLocalizationBatch(
            session.GetLocalizationReviewRows(),
            {.kind = RedConfigure::Workflow::LocalizationBatchKind::FindReplace,
             .targetCulture = L"cs-CZ",
             .findText = L"upraveno",
             .replaceText = L"batch",
             .rowIndices = {rowIndex}});
        ok = Require(session.ApplyLocalizationBatch(readyBatch) == RedConfigure::Workflow::BatchApprovalResult::Applied,
                     L"A fresh localization batch preview should apply as one edit.") &&
             ok;
        ok = Require(session.Undo(), L"A successful localization batch should be exactly one Undo step.") && ok;

        RedConfigure::Ui::LocalizationPagePresenter localizationPresenter;
        const RedConfigure::Workflow::LocalizationBatchRequest presenterRequest{
            .kind = RedConfigure::Workflow::LocalizationBatchKind::FindReplace,
            .targetCulture = L"cs-CZ",
            .findText = L"upraveno",
            .replaceText = L"first",
            .rowIndices = {rowIndex}};
        const auto firstPresentation = localizationPresenter.Execute(session, presenterRequest);
        RedConfigure::Workflow::LocalizationBatchRequest changedPresenterRequest = presenterRequest;
        changedPresenterRequest.replaceText = L"second";
        const auto changedPresentation = localizationPresenter.Execute(session, changedPresenterRequest);
        ok = Require(firstPresentation.phase == RedConfigure::Ui::BatchInteractionPhase::Preview &&
                         changedPresentation.phase == RedConfigure::Ui::BatchInteractionPhase::Preview,
                     L"Changing any localization request argument should replace approval with a new preview instead of applying.") &&
             ok;
        const auto presenterApply = localizationPresenter.Execute(session, changedPresenterRequest);
        ok = Require(presenterApply.phase == RedConfigure::Ui::BatchInteractionPhase::Apply &&
                         presenterApply.result == RedConfigure::Workflow::BatchApprovalResult::Applied && session.Undo(),
                     L"An unchanged localization presenter request should apply once and remain one Undo step.") &&
             ok;

        const auto* updatedRow  = FindReviewRow(session.GetLocalizationReviewRows(), L"App", L"IDS_HELLO");
        const auto* updatedCell = updatedRow ? FindReviewTargetCell(*updatedRow, L"cs-CZ") : nullptr;
        ok                      = Require(updatedCell != nullptr && updatedCell->targetText == L"Ahoj upraveno {0}" && updatedCell->dirty,
                                          L"Rejected review target edits should keep the last valid target text and dirty state.") &&
                                  ok;
    }

    std::vector<RedConfigure::LocalizationExportPreview> previews;
    ok = Require(SUCCEEDED(session.BuildLocalizationReviewExportPreviews(previews)), L"Review export previews should build.") && ok;
    ok = Require(previews.size() == 1u, L"Review export previews should include only changed owner/culture satellite files.") && ok;
    ok = Require(ContainsExportPreview(previews, L"App", L"cs-CZ", L"Ahoj upraveno {0}"),
                 L"Review export previews should contain edited App cs-CZ target text.") &&
         ok;
    ok = Require(! ContainsExportPreview(previews, L"Plugin", L"fr-FR", L"Plugin seulement"),
                 L"Review export previews should not include unchanged Plugin fr-FR target text.") &&
         ok;

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestLocalizationReviewSkipsBadResourceFilesAndKeepsWorkspaceOpen()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureLocalizationReviewBadRcTest", ec);
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for bad localization resource tests.");
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
                      L"Failed to write healthy app project fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"App.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Hello {0}"
END
)rc"),
                      L"Failed to write healthy app rc fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"Lang" / L"fr-FR" / L"App-fr-FR.rc",
                                        R"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Bonjour {0}"
END
)rc"),
                      L"Failed to write healthy app target rc fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Broken" / L"Broken.vcxproj",
                                        R"xml(<?xml version="1.0" encoding="utf-8"?>
<Project>
  <ItemGroup>
    <ResourceCompile Include="Broken.rc" />
  </ItemGroup>
</Project>)xml"),
                      L"Failed to write broken project fixture.") &&
              ok;
    constexpr std::array<uint8_t, 3u> invalidUtf16Le{{0xFFu, 0xFEu, 0x41u}};
    ok = Require(WriteTestBinaryFile(tempRoot / L"Broken" / L"Broken.rc", invalidUtf16Le), L"Failed to write invalid UTF-16 resource fixture.") && ok;

    RedConfigure::RedConfigureSession session;
    ok = Require(SUCCEEDED(session.LoadWorkspace(tempRoot, L"fr-FR")), L"Session should keep loading when one resource file is unreadable.") && ok;
    ok = Require(! session.GetWorkspace().errors.empty(), L"Session should report unreadable resource files in workspace errors.") && ok;
    bool mentionedBrokenPath = false;
    for (const std::wstring& error : session.GetWorkspace().errors)
    {
        mentionedBrokenPath = mentionedBrokenPath || error.find(L"Broken.rc") != std::wstring::npos;
    }
    ok = Require(mentionedBrokenPath, L"Workspace errors should identify the bad resource path.") && ok;
    const RedConfigure::Workflow::ValidationSummary typedValidation = session.Validate();
    ok = Require(std::ranges::any_of(typedValidation.issues,
                                     [](const RedConfigure::Workflow::ValidationIssue& issue) noexcept
                                     {
                                         return issue.category == RedConfigure::Workflow::ValidationCategory::Workspace &&
                                                issue.code == RedConfigure::Workflow::ValidationCode::WorkspaceProcessingError &&
                                                issue.severity == RedConfigure::Workflow::ValidationSeverity::Error;
                                     }),
                 L"Workspace validation should classify source failures by typed category, code, and severity.") &&
         ok;

    ok = Require(FindReviewRow(session.GetLocalizationReviewRows(), L"App", L"IDS_HELLO") != nullptr,
                 L"Review rows from healthy owners should still load when another owner is bad.") &&
         ok;
    ok = Require(! session.GetInventoryEntries().empty(), L"Active-owner inventory should still populate from the healthy owner.") && ok;

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestLocalizationReviewGridModelShowsOwnersEnglishAndVisibleCultures()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureLocalizationReviewGridModelTest", ec);
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for localization review grid model tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    bool ok = WriteLocalizationReviewWorkspaceFixture(tempRoot);

    RedConfigure::RedConfigureSession session;
    ok = Require(SUCCEEDED(session.LoadWorkspace(tempRoot, L"fr-FR")), L"Session should load the localization review grid model fixture.") && ok;

    RedConfigure::LocalizationReviewViewOptions options;
    options.visibleOwnerNames          = {L"App", L"Plugin"};
    options.visibleCultureNames        = {L"fr-FR", L"cs-CZ"};
    const std::vector<size_t> viewRows = RedConfigure::BuildLocalizationReviewView(session.GetLocalizationReviewRows(), options);

    RedConfigure::Ui::LocalizationReviewGridModel model(nullptr, session);
    model.SetViewRows(viewRows);
    model.SetVisibleCultures(options.visibleCultureNames);

    ok = Require(model.GetRowCount() == viewRows.size(), L"Review grid model should expose projected rows.") && ok;
    ok = Require(model.GetColumnCount() == 6u, L"Review grid model should expose owner, ID, English, visible cultures, and status columns.") && ok;
    ok = Require(model.GetColumn(0u).id == L"owner", L"Review grid model first column should be owner.") && ok;
    ok = Require(model.GetColumn(2u).id == L"english", L"Review grid model should expose English as a read-only source column.") && ok;
    ok = Require(model.GetColumn(2u).multiline, L"Review grid model English column should allow wrapped two-line rows.") && ok;
    ok = Require(model.GetColumn(3u).id == L"target:fr-FR", L"Review grid model should expose the first visible culture column.") && ok;
    ok = Require(model.GetColumn(3u).multiline, L"Review grid model target columns should allow wrapped two-line rows.") && ok;
    ok = Require(model.GetColumn(4u).id == L"target:cs-CZ", L"Review grid model should expose the second visible culture column.") && ok;

    RedSalamander::DxUi::GridCellData cell;
    model.GetCellData(0u, 0u, cell);
    ok = Require(cell.text == L"App", L"Review grid model owner cell should show owner name.") && ok;
    model.GetCellData(0u, 2u, cell);
    ok = Require(cell.text == L"Hello {0}", L"Review grid model English cell should show source text.") && ok;
    model.GetCellData(0u, 3u, cell);
    ok = Require(cell.text == L"Bonjour {0}", L"Review grid model target culture cell should show target text.") && ok;
    ok = Require(cell.multiline, L"Review grid model target cells should render as multiline content.") && ok;
    model.GetCellData(0u, 5u, cell);
    ok = Require(cell.text.empty(), L"Review grid model should leave clean status cells blank instead of showing OK.") && ok;

    bool foundWarningRow = false;
    for (size_t rowIndex = 0u; rowIndex < model.GetRowCount(); ++rowIndex)
    {
        foundWarningRow = foundWarningRow || model.GetRowStyle(rowIndex).tone == RedSalamander::DxUi::GridRowTone::Warning;
    }
    ok = Require(foundWarningRow, L"Review grid model should mark rows with visible target problems as warnings.") && ok;
    ok = Require(model.GetStableRowId(0u) != 1u && model.GetStableRowId(0u) == model.GetStableRowId(0u),
                 L"Review grid model stable row IDs should be deterministic owner/ID hashes, not view offsets.") &&
         ok;
    ok = Require(model.FindRowByStableId(model.GetStableRowId(1u)).value_or(0u) == 1u,
                 L"Review grid model should resolve stable row IDs within the projected rows.") &&
         ok;
    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestLocalizationReviewSurfacesMissingTranslations()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureLocalizationReviewMissingTest", ec);
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for missing-translation review tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    bool ok = WriteLocalizationReviewWorkspaceFixture(tempRoot);

    RedConfigure::RedConfigureSession session;
    ok = Require(SUCCEEDED(session.LoadWorkspace(tempRoot, L"fr-FR")), L"Session should load the missing-translation review fixture.") && ok;

    const auto rows       = session.GetLocalizationReviewRows();
    const auto rowMatches = [&rows](size_t rowIndex, std::wstring_view ownerName, std::wstring_view id) noexcept
    { return rowIndex < rows.size() && rows[rowIndex].ownerName == ownerName && rows[rowIndex].id == id; };

    RedConfigure::LocalizationReviewViewOptions options;
    options.visibleOwnerNames             = {L"App", L"Plugin"};
    options.visibleCultureNames           = {L"fr-FR", L"cs-CZ"};
    options.statusFilter                  = RedConfigure::LocalizationStatusFilter::Problems;
    const std::vector<size_t> problemView = RedConfigure::BuildLocalizationReviewView(rows, options);

    bool problemsIncludeMissingRow = false;
    for (const size_t rowIndex : problemView)
    {
        problemsIncludeMissingRow = problemsIncludeMissingRow || rowMatches(rowIndex, L"App", L"IDS_APP_ONLY");
    }
    ok = Require(problemsIncludeMissingRow, L"Problems filter should include rows whose visible cultures have no existing translation.") && ok;

    options.statusFilter             = RedConfigure::LocalizationStatusFilter::Ok;
    const std::vector<size_t> okView = RedConfigure::BuildLocalizationReviewView(rows, options);
    for (const size_t rowIndex : okView)
    {
        ok = Require(! rowMatches(rowIndex, L"App", L"IDS_APP_ONLY"), L"OK filter should exclude rows with missing translations.") && ok;
    }

    options.statusFilter              = RedConfigure::LocalizationStatusFilter::All;
    const std::vector<size_t> allView = RedConfigure::BuildLocalizationReviewView(rows, options);

    RedConfigure::Ui::LocalizationReviewGridModel model(nullptr, session);
    model.SetViewRows(allView);
    model.SetVisibleCultures(options.visibleCultureNames);

    std::optional<size_t> missingViewRow;
    std::optional<size_t> warningViewRow;
    std::optional<size_t> cleanViewRow;
    for (size_t viewRow = 0u; viewRow < allView.size(); ++viewRow)
    {
        if (rowMatches(allView[viewRow], L"App", L"IDS_APP_ONLY"))
        {
            missingViewRow = viewRow;
        }
        else if (rowMatches(allView[viewRow], L"Plugin", L"IDS_HELLO"))
        {
            warningViewRow = viewRow;
        }
        else if (rowMatches(allView[viewRow], L"App", L"IDS_HELLO"))
        {
            cleanViewRow = viewRow;
        }
    }

    ok = Require(missingViewRow.has_value() && model.GetRowStyle(missingViewRow.value()).tone == RedSalamander::DxUi::GridRowTone::Info,
                 L"Missing-translation rows should render the distinct Missing row tone.") &&
         ok;
    ok = Require(warningViewRow.has_value() && model.GetRowStyle(warningViewRow.value()).tone == RedSalamander::DxUi::GridRowTone::Warning,
                 L"Placeholder problems should keep warning-tone precedence over the Missing tone.") &&
         ok;
    ok = Require(cleanViewRow.has_value() && model.GetRowStyle(cleanViewRow.value()).tone == RedSalamander::DxUi::GridRowTone::None,
                 L"Fully translated rows should keep the default row tone.") &&
         ok;

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestLocalizationReviewCanCreateAndExportNewCulture()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureLocalizationReviewNewCultureTest", ec);
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for localization review new-culture tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    bool ok = WriteLocalizationReviewWorkspaceFixture(tempRoot);

    RedConfigure::RedConfigureSession session;
    ok = Require(SUCCEEDED(session.LoadWorkspace(tempRoot, L"fr-FR")), L"Session should load the localization review new-culture fixture.") && ok;
    ok = Require(session.EnsureLocalizationReviewCulture(L"de-DE"), L"Session should create an in-memory target culture for review.") && ok;
    ok = Require(ContainsCulture(session.GetLocalizationReviewCultures(), L"de-DE"), L"New target culture should be listed with review cultures.") && ok;

    const auto* appHello = FindReviewRow(session.GetLocalizationReviewRows(), L"App", L"IDS_HELLO");
    ok                   = Require(appHello != nullptr, L"New-culture fixture should expose App IDS_HELLO.") && ok;
    if (appHello)
    {
        const size_t rowIndex = static_cast<size_t>(appHello - session.GetLocalizationReviewRows().data());
        ok = Require(session.UpdateLocalizationReviewTarget(rowIndex, L"de-DE", L"Hallo {0}"), L"Review target edit should accept a newly created culture.") &&
             ok;
    }

    std::vector<RedConfigure::LocalizationExportPreview> previews;
    ok = Require(SUCCEEDED(session.BuildLocalizationReviewExportPreviews(previews)), L"Review export previews should include new target cultures.") && ok;
    ok = Require(previews.size() == 1u, L"New-culture preview should include only owner/culture files with edits.") && ok;
    ok = Require(ContainsExportPreview(previews, L"App", L"de-DE", L"Hallo {0}"), L"New target culture preview should contain edited text.") && ok;

    const std::filesystem::path outputRoot = tempRoot / L"ReviewOut";
    ok = Require(SUCCEEDED(session.ExportLocalizationReview(outputRoot)), L"Session should write review satellite rc files.") && ok;
    ok = Require(std::filesystem::exists(outputRoot / L"App-de-DE.rc", ec), L"Review export should write the new App de-DE rc file.") && ok;
    ok = Require(! std::filesystem::exists(outputRoot / L"Plugin-de-DE.rc", ec), L"Review export should not write unchanged owner/culture files.") && ok;

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestLocalizationReviewProjectionPerformanceMetric()
{
    constexpr size_t ownerCount      = 8u;
    constexpr size_t stringsPerOwner = 600u;
    const std::vector<std::wstring> cultures{L"fr-FR", L"cs-CZ", L"ja-JP", L"sk-SK"};

    std::vector<RedConfigure::LocalizationReviewRow> rows;
    rows.reserve(ownerCount * stringsPerOwner);
    RedConfigure::LocalizationReviewViewOptions options;
    options.visibleOwnerNames.reserve(ownerCount);
    options.visibleCultureNames = cultures;

    for (size_t ownerIndex = 0u; ownerIndex < ownerCount; ++ownerIndex)
    {
        const std::wstring ownerName = L"PerfOwner" + std::to_wstring(ownerIndex);
        options.visibleOwnerNames.push_back(ownerName);
        for (size_t stringIndex = 0u; stringIndex < stringsPerOwner; ++stringIndex)
        {
            RedConfigure::LocalizationReviewRow row;
            row.ownerName  = ownerName;
            row.id         = L"IDS_PERF_" + std::to_wstring(ownerIndex) + L"_" + std::to_wstring(stringIndex);
            row.sourceText = L"Source " + std::to_wstring(ownerIndex) + L" " + std::to_wstring(stringIndex) + L" {0}";
            row.targets.reserve(cultures.size());
            for (const std::wstring& culture : cultures)
            {
                RedConfigure::LocalizationTargetCell cell;
                cell.cultureName            = culture;
                cell.targetText             = culture + L" target " + std::to_wstring(ownerIndex) + L" " + std::to_wstring(stringIndex) + L" {0}";
                cell.validation             = RedConfigure::Localization::ValidatePlaceholders(row.sourceText, cell.targetText);
                cell.hasExistingTranslation = true;
                row.targets.push_back(std::move(cell));
            }
            rows.push_back(std::move(row));
        }
    }

    const auto started             = std::chrono::steady_clock::now();
    const std::vector<size_t> view = RedConfigure::BuildLocalizationReviewView(rows, options);
    const auto finished            = std::chrono::steady_clock::now();
    const auto elapsedMicros       = std::chrono::duration_cast<std::chrono::microseconds>(finished - started).count();

    std::wcout << L"redconfigure.localization.review.all_owners_all_languages rows=" << rows.size() << L" visibleRows=" << view.size() << L" cultures="
               << cultures.size() << L" elapsedMicros=" << elapsedMicros << L'\n';

    return Require(view.size() == rows.size(), L"Localization review perf projection should keep every checked owner/language row visible.");
}

[[nodiscard]] bool TestRedConfigureRepoSizedScanAndValidationPerformance()
{
    constexpr size_t ownerCount = 6u;
    constexpr size_t stringsPerOwner = 250u;
    const std::array<std::wstring, 4> cultures{{L"fr-FR", L"cs-CZ", L"ja-JP", L"sk-SK"}};
    std::error_code ec;
    const std::filesystem::path root = AcquireRedConfigureTestSandbox(L"RepoSizedScanValidationPerf", ec);
    if (ec) return Require(false, L"Could not create repo-sized RedConfigure performance sandbox.");
    std::filesystem::remove_all(root, ec);

    bool ok = true;
    for (size_t ownerIndex = 0u; ownerIndex < ownerCount; ++ownerIndex)
    {
        const std::wstring ownerName = L"PerfOwner" + std::to_wstring(ownerIndex);
        const std::filesystem::path ownerRoot = root / ownerName;
        const std::string project = std::format("<Project><ItemGroup><ResourceCompile Include=\"{}.rc\" /></ItemGroup></Project>\n", ownerIndex);
        ok = Require(WriteTestTextFile(ownerRoot / (ownerName + L".vcxproj"), project), L"Could not write repo-sized perf project fixture.") && ok;

        std::string source = "#include \"resource.h\"\nSTRINGTABLE\nBEGIN\n";
        for (size_t stringIndex = 0u; stringIndex < stringsPerOwner; ++stringIndex)
        {
            source += std::format(" IDS_PERF_{}_{} \"Source {} {} {{0}}\"\n", ownerIndex, stringIndex, ownerIndex, stringIndex);
        }
        source += "END\n";
        ok = Require(WriteTestTextFile(ownerRoot / (std::to_wstring(ownerIndex) + L".rc"), source), L"Could not write repo-sized source RC fixture.") && ok;

        for (const std::wstring& culture : cultures)
        {
            std::string target = "#include \"resource.h\"\nSTRINGTABLE\nBEGIN\n";
            for (size_t stringIndex = 0u; stringIndex < stringsPerOwner; ++stringIndex)
            {
                target += std::format(" IDS_PERF_{}_{} \"Target {} {} {{0}}\"\n", ownerIndex, stringIndex, ownerIndex, stringIndex);
            }
            target += "END\n";
            ok = Require(WriteTestTextFile(ownerRoot / L"Lang" / culture / (ownerName + L"-" + culture + L".rc"), target),
                         L"Could not write repo-sized target RC fixture.") &&
                 ok;
        }
    }
    for (size_t themeIndex = 0u; themeIndex < 10u; ++themeIndex)
    {
        const std::string theme = std::format(
            "{{ formatVersion: 2, id: \"user/perf-{}\", name: \"Perf {}\", baseThemeId: \"builtin/dark\", colors: {{ \"app.accent\": \"#336699\" }} }}\n",
            themeIndex,
            themeIndex);
        ok = Require(WriteTestTextFile(root / L"Specs" / L"Themes" / (L"Perf" + std::to_wstring(themeIndex) + L".theme.json5"), theme),
                     L"Could not write repo-sized theme fixture.") &&
             ok;
    }
    if (! ok) return false;

    RedConfigure::RedConfigureSession session;
    const auto scanStarted = std::chrono::steady_clock::now();
    const HRESULT loadHr = session.LoadWorkspace(root, L"fr-FR");
    const auto scanElapsed = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - scanStarted);
    ok = Require(SUCCEEDED(loadHr), L"Repo-sized RedConfigure performance workspace should load.") && ok;
    ok = Require(session.GetLocalizationReviewRows().size() == ownerCount * stringsPerOwner,
                 L"Repo-sized RedConfigure performance workspace should expose every resource row.") &&
         ok;

    const auto validationStarted = std::chrono::steady_clock::now();
    const RedConfigure::Workflow::ValidationSummary validation = session.Validate();
    const auto validationElapsed = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - validationStarted);
    ok = Require(validation.errorCount == 0u, L"Repo-sized RedConfigure performance fixture should validate without errors.") && ok;

    const auto previewStarted = std::chrono::steady_clock::now();
    ok = Require(session.UpdateThemeColor(L"app.accent", L"#5577AA"), L"Repo-sized RedConfigure preview update should succeed.") && ok;
    const auto previewElapsed = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - previewStarted);

    std::wcout << L"redconfigure.repo_sized owners=" << ownerCount << L" rows=" << session.GetLocalizationReviewRows().size() << L" themes="
               << session.GetThemeCatalog().themes.size() << L" scanMs=" << scanElapsed.count() << L" validationMs=" << validationElapsed.count()
               << L" previewUs=" << previewElapsed.count() << L'\n';
    ok = Require(scanElapsed < std::chrono::milliseconds(500), L"Repo-sized RedConfigure scan should stay under the 500 ms target.") && ok;
    ok = Require(validationElapsed < std::chrono::milliseconds(250), L"Repo-sized RedConfigure validation should stay under the 250 ms target.") && ok;
    ok = Require(previewElapsed < std::chrono::milliseconds(16), L"Theme preview model update should stay within one 16 ms frame budget.") && ok;
    std::filesystem::remove_all(root, ec);
    return ok;
}

[[nodiscard]] bool TestSplashScreenCloseGuardSuppressesPendingOpen()
{
    bool ok = true;

    wil::unique_handle closeEvent(::CreateEventW(nullptr, TRUE, TRUE, nullptr));
    ok = Require(closeEvent.get() != nullptr, L"Failed to create close event for RedConfigure splash guard test.") && ok;
    ok = Require(RedConfigure::SplashScreen::Detail::ShouldAbortPendingOpen(std::stop_token{}, closeEvent.get()),
                 L"Signaled close event should suppress RedConfigure splash open before the worker creates a window.") &&
         ok;

    std::stop_source stopSource;
    stopSource.request_stop();
    ok = Require(RedConfigure::SplashScreen::Detail::ShouldAbortPendingOpen(stopSource.get_token(), nullptr),
                 L"Requested stop token should suppress RedConfigure splash open before the worker creates a window.") &&
         ok;
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
    ok                = Require(merged.size() == 2u, L"Merged resource model should contain all source strings.") && ok;
    ok                = Require(merged[1].targetText == L"Beta FR {0}", L"Merged resource model should keep target translations by ID.") && ok;

    const std::wstring output = RedConfigure::Localization::BuildSatelliteRcStringTable(L"resource.h", L"fr-FR", merged);
    ok                        = Require(output.find(L"#include \"resource.h\"") != std::wstring::npos, L"Satellite writer should include resource.h.") && ok;
    ok                        = Require(output.find(L"IDS_ALPHA") < output.find(L"IDS_BETA"), L"Satellite writer should sort entries by ID.") && ok;
    ok = Require(output.find(L"\"Beta FR {0}\"") != std::wstring::npos, L"Satellite writer should keep positional placeholders unchanged.") && ok;
    return ok;
}

[[nodiscard]] bool TestThemeCatalogLoadsThemeFiles()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureThemeCatalogTest", ec);
    if (ec)
    {
        return Require(false, L"Could not resolve a temporary directory for theme catalog tests.");
    }

    std::filesystem::remove_all(tempRoot, ec);
    bool ok = true;
    ok      = Require(WriteTestTextFile(tempRoot / L"Good.theme.json5",
                                        R"json5({
  formatVersion: 2,
  id: "user/catalog-good",
  name: "Catalog Good",
  baseThemeId: "builtin/dark",
  colors: {
    "app.accent": "#336699",
  },
})json5"),
                      L"Failed to write valid theme catalog fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Bad.theme.json5", "{ id: 'user/bad' }\n"), L"Failed to write invalid theme catalog fixture.") && ok;

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
    ok      = Require(model.TryEditOverride(L"folderView.itemBackgroundSelected", L"#336699"), L"Theme preview should accept valid color edits.") && ok;
    ok      = Require(model.GetEffectiveColor(L"folderView.itemBackgroundSelected").value_or(0u) == 0xFF336699u,
                      L"Theme preview should update immediately after valid color edits.") &&
              ok;
    ok      = Require(! model.TryEditOverride(L"folderView.itemBackgroundSelected", L"not-a-color"), L"Theme preview should reject invalid color text.") && ok;
    ok      = Require(model.GetEffectiveColor(L"folderView.itemBackgroundSelected").value_or(0u) == 0xFF336699u,
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
    ok = Require(model.TryEditOverride(L"folderView.itemBackgroundSelected", L"darken(app.accent,25%)"), L"Theme preview should accept darken expressions.") &&
         ok;
    ok = Require(model.GetEffectiveColor(L"folderView.itemBackgroundSelected").value_or(0u) == 0xFF606060u,
                 L"Darken expressions should derive the expected color.") &&
         ok;

    ok =
        Require(model.TryEditOverride(L"folderView.background", L"blend(menu.background,app.accent,50%)"), L"Theme preview should accept blend expressions.") &&
        ok;
    ok = Require(model.GetEffectiveColor(L"folderView.background").value_or(0u) == 0xFF505050u, L"Blend expressions should derive the expected color.") && ok;

    ok = Require(model.TryEditOverride(L"window.background", L"ref(folderView.background)"), L"Theme preview should accept reference expressions.") && ok;
    ok = Require(model.GetEffectiveColor(L"window.background").value_or(0u) == 0xFF505050u,
                 L"Reference expressions should reuse the referenced effective color.") &&
         ok;

    ok = Require(! model.TryEditOverride(L"app.accent", L"ref(folderView.itemBackgroundSelected)"), L"Theme preview should reject color expression cycles.") &&
         ok;
    ok = Require(model.GetEffectiveColor(L"app.accent").value_or(0u) == 0xFF808080u, L"Rejected expression cycles should keep the previous valid color.") && ok;
    return ok;
}

[[nodiscard]] std::optional<std::filesystem::path> FindWindowsResourceCompiler()
{
    const std::wstring programFilesX86 = GetEnvironmentString(L"ProgramFiles(x86)");
    if (programFilesX86.empty())
    {
        return std::nullopt;
    }

    const std::filesystem::path binRoot = std::filesystem::path(programFilesX86) / L"Windows Kits" / L"10" / L"bin";
    std::error_code ec;
    std::vector<std::filesystem::path> candidates;
    for (std::filesystem::directory_iterator it(binRoot, ec), end; ! ec && it != end; it.increment(ec))
    {
        if (! it->is_directory(ec))
        {
            ec.clear();
            continue;
        }
        const std::filesystem::path candidate = it->path() / L"x64" / L"rc.exe";
        if (std::filesystem::is_regular_file(candidate, ec))
        {
            candidates.push_back(candidate);
        }
        ec.clear();
    }
    if (candidates.empty())
    {
        return std::nullopt;
    }
    std::ranges::sort(candidates);
    return candidates.back();
}

[[nodiscard]] bool TestGeneratedRcCompilesWhenWindowsSdkIsAvailable()
{
    const std::optional<std::filesystem::path> rcExe = FindWindowsResourceCompiler();
    if (! rcExe.has_value())
    {
        std::wcout << L"RedConfigureTests: Windows SDK rc.exe unavailable; generated RC compiler validation skipped.\n";
        return true;
    }

    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"GeneratedRcCompilerValidation", ec);
    if (ec)
    {
        return Require(false, L"Could not create the generated RC compiler-validation sandbox.");
    }
    std::filesystem::remove_all(tempRoot, ec);
    std::filesystem::create_directories(tempRoot, ec);
    if (ec)
    {
        return Require(false, L"Could not reset the generated RC compiler-validation sandbox.");
    }

    const std::vector<RedConfigure::Localization::MergedStringEntry> merged = {
        {.id = L"IDS_ALPHA", .sourceText = L"Alpha", .targetText = L"Alpha traduit"},
        {.id = L"IDS_FORMATTED", .sourceText = L"Value {0}", .targetText = L"Valeur {0}"},
    };
    const std::wstring generated = RedConfigure::Localization::BuildSatelliteRcStringTable(L"resource.h", L"fr-FR", merged);
    bool ok = Require(WriteTestUtf16LeTextFile(tempRoot / L"generated.rc", generated, true), L"Could not write generated RC compiler fixture.");
    ok = Require(WriteTestTextFile(tempRoot / L"resource.h", "#define IDS_ALPHA 1001\n#define IDS_FORMATTED 1002\n"),
                 L"Could not write generated RC resource header fixture.") &&
         ok;
    if (! ok)
    {
        return false;
    }

    const std::filesystem::path outputPath = tempRoot / L"generated.res";
    std::wstring commandLine = std::format(L"\"{}\" /nologo /fo \"{}\" /I \"{}\" \"{}\"",
                                           rcExe->wstring(),
                                           outputPath.wstring(),
                                           tempRoot.wstring(),
                                           (tempRoot / L"generated.rc").wstring());
    STARTUPINFOW startupInfo{.cb = sizeof(STARTUPINFOW)};
    PROCESS_INFORMATION processInfo{};
    if (! CreateProcessW(rcExe->c_str(), commandLine.data(), nullptr, nullptr, FALSE, CREATE_NO_WINDOW, nullptr, tempRoot.c_str(), &startupInfo, &processInfo))
    {
        return Require(false, L"Could not start Windows SDK rc.exe for generated RC validation.");
    }
    wil::unique_handle process(processInfo.hProcess);
    wil::unique_handle thread(processInfo.hThread);
    if (WaitForSingleObject(process.get(), 30'000u) != WAIT_OBJECT_0)
    {
        static_cast<void>(TerminateProcess(process.get(), ERROR_TIMEOUT));
        return Require(false, L"Windows SDK rc.exe timed out while validating generated RC output.");
    }
    DWORD exitCode = ERROR_GEN_FAILURE;
    ok = Require(GetExitCodeProcess(process.get(), &exitCode) != FALSE && exitCode == 0u,
                 L"Windows SDK rc.exe rejected RedConfigure generated RC output.") &&
         ok;
    ok = Require(std::filesystem::is_regular_file(outputPath, ec) && ! ec, L"Windows SDK rc.exe did not create the expected .res file.") && ok;
    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestRedConfigureWorkflowModels()
{
    bool ok = true;
    RedConfigure::RedConfigureSession emptySession;
    const RedConfigure::Ui::StartPageSummary emptySummary = RedConfigure::Ui::StartPagePresenter::Build(emptySession);
    ok = Require(emptySummary.resourceOwnerCount == 0u && emptySummary.themeFileCount == 0u && emptySummary.scanErrorCount == 0u,
                 L"The start-page presenter should project session counts without constructing the workbench root.") &&
         ok;
    ok = Require(RedConfigure::Ui::ThemesPagePresenter::GetOriginResourceId(RedConfigure::Themes::ThemeCatalogOrigin::User) ==
                     IDS_REDCONFIGURE_THEME_ORIGIN_USER &&
                     RedConfigure::Ui::ReviewExportPagePresenter::GetCategoryResourceId(
                         RedConfigure::Workflow::ValidationCategory::Localization) == IDS_REDCONFIGURE_CATEGORY_LOCALIZATION &&
                     RedConfigure::Ui::ReviewExportPagePresenter::GetMessageResourceId(
                         RedConfigure::Workflow::ValidationCode::PlaceholderMismatch) == IDS_REDCONFIGURE_VALIDATION_PLACEHOLDER_MISMATCH,
                 L"Page presenters should route typed origin and validation policy to stable resource identifiers.") &&
         ok;
    RedConfigure::Workflow::LanguageColumnModel columns;
    const std::vector<std::wstring> initialCultures{L"fr-FR", L"de-DE", L"ja-JP"};
    columns.Set(initialCultures);
    ok = Require(columns.SetPinned(L"ja-JP", true), L"Language columns should support pinning.") && ok;
    ok = Require(columns.Move(L"de-DE", 1u), L"Language columns should support reordering.") && ok;
    ok = Require(columns.Remove(L"fr-FR"), L"Language columns should support removal.") && ok;
    ok = Require(columns.Add(L"cs-CZ"), L"Language columns should support addition.") && ok;
    const std::vector<std::wstring> ordered = columns.GetOrderedCultures();
    ok = Require(ordered.size() == 3u && ordered.front() == L"ja-JP" && ordered.back() == L"cs-CZ",
                 L"Pinned language columns should remain first while preserving configured order.") &&
         ok;

    RedConfigure::LocalizationReviewRow row;
    row.ownerName = L"MenuOwner";
    row.id = L"IDS_MENU_OPEN";
    row.sourceText = L"&Open {0}";
    row.targets.push_back({.cultureName = L"fr-FR", .targetText = L"&Ouvrir {0}", .hasExistingTranslation = true});
    row.targets.push_back({.cultureName = L"de-DE", .targetText = L"Openen { 0 }", .hasExistingTranslation = true});
    std::vector<RedConfigure::LocalizationReviewRow> rows{row};
    struct LocalizationRecipeCase
    {
        RedConfigure::Workflow::LocalizationBatchKind kind;
        std::wstring sourceCulture;
        std::wstring findText;
        std::wstring replaceText;
        std::wstring expected;
        bool expectedReviewed = false;
    };
    const std::array localizationCases = {
        LocalizationRecipeCase{.kind = RedConfigure::Workflow::LocalizationBatchKind::CopyEnglish, .expected = L"&Open {0}"},
        LocalizationRecipeCase{.kind = RedConfigure::Workflow::LocalizationBatchKind::CopyCulture,
                               .sourceCulture = L"fr-FR",
                               .expected = L"&Ouvrir {0}"},
        LocalizationRecipeCase{.kind = RedConfigure::Workflow::LocalizationBatchKind::Clear, .expected = L""},
        LocalizationRecipeCase{.kind = RedConfigure::Workflow::LocalizationBatchKind::FindReplace,
                               .findText = L"Openen",
                               .replaceText = L"Neu",
                               .expected = L"Neu { 0 }"},
        LocalizationRecipeCase{.kind = RedConfigure::Workflow::LocalizationBatchKind::NormalizePlaceholderWhitespace,
                               .expected = L"Openen {0}"},
        LocalizationRecipeCase{.kind = RedConfigure::Workflow::LocalizationBatchKind::PreserveAccelerators, .expected = L"&Openen { 0 }"},
        LocalizationRecipeCase{.kind = RedConfigure::Workflow::LocalizationBatchKind::MarkReviewed,
                               .expected = L"Openen { 0 }",
                               .expectedReviewed = true},
    };
    for (const LocalizationRecipeCase& testCase : localizationCases)
    {
        const RedConfigure::Workflow::LocalizationBatchRequest request{.kind = testCase.kind,
                                                                        .sourceCulture = testCase.sourceCulture,
                                                                        .targetCulture = L"de-DE",
                                                                        .findText = testCase.findText,
                                                                        .replaceText = testCase.replaceText};
        const auto preview = RedConfigure::Workflow::PreviewLocalizationBatch(rows, request);
        ok = Require(preview.result == RedConfigure::Workflow::BatchApprovalResult::Ready && preview.changes.size() == 1u &&
                         preview.changes.front().after == testCase.expected &&
                         preview.changes.front().afterReviewed == testCase.expectedReviewed,
                     L"Every localization batch kind should produce the characterized typed preview.") &&
             ok;
    }
    const auto invalidLocalization = RedConfigure::Workflow::PreviewLocalizationBatch(
        rows,
        {.kind = RedConfigure::Workflow::LocalizationBatchKind::FindReplace, .targetCulture = L"de-DE"});
    ok = Require(invalidLocalization.result == RedConfigure::Workflow::BatchApprovalResult::Invalid,
                 L"A malformed localization batch request should be typed as invalid.") &&
         ok;

    const auto clipboard = RedConfigure::Workflow::ParseClipboardMatrix(L"A\tB\r\nC\tD");
    ok = Require(clipboard.rows.size() == 2u && clipboard.rows.front().size() == 2u &&
                     RedConfigure::Workflow::SerializeClipboardMatrix(clipboard) == L"A\tB\r\nC\tD",
                 L"Rectangular localization clipboard data should round trip as TSV.") &&
         ok;
    const auto trailingClipboard = RedConfigure::Workflow::ParseClipboardMatrix(L"A\tB\r\n");
    ok = Require(trailingClipboard.rows.size() == 1u && trailingClipboard.rows.front().size() == 2u &&
                     RedConfigure::Workflow::SerializeClipboardMatrix(trailingClipboard) == L"A\tB",
                 L"Clipboard parser should ignore a terminal newline instead of synthesizing an empty trailing row.") &&
         ok;
    const auto emptyClipboard = RedConfigure::Workflow::ParseClipboardMatrix(L"");
    ok = Require(emptyClipboard.rows.empty(), L"Empty clipboard text should not synthesize an editable row.") && ok;

    rows.push_back(rows.front());
    rows.back().id = L"IDS_MENU_OTHER";
    const auto duplicates = RedConfigure::Workflow::FindDuplicateSiblingAccelerators(rows, L"fr-FR");
    ok = Require(duplicates.size() == 1u && duplicates.front().resourceIds.size() == 2u,
                 L"Sibling menu accelerator validation should report duplicates.") &&
         ok;

    Common::Settings::ThemeDefinition theme;
    theme.id = L"user/workflow";
    theme.name = L"Workflow";
    theme.baseThemeId = L"builtin/dark";
    theme.palette.emplace(L"accent", 0xFF336699u);
    theme.palette.emplace(L"alternate", 0xFF663399u);
    theme.palette.emplace(L"warning", 0xFFFFAA00u);
    theme.palette.emplace(L"error", 0xFFFF0000u);
    theme.palette.emplace(L"success", 0xFF00AA00u);
    theme.colors.emplace(L"menu.background", 0xFF000000u);
    Common::Settings::ThemeColorSource menuText;
    ok = Require(SUCCEEDED(Common::Settings::ParseThemeColorSource(L"blend(palette.accent,palette.alternate,25%)", menuText)),
                 L"Theme workflow reference fixture should parse.") &&
         ok;
    theme.colors.emplace(L"menu.text", std::move(menuText));
    RedConfigure::Themes::ThemePreviewModel model;
    model.SetTheme(theme);
    const auto metadata = RedConfigure::Workflow::BuildThemeTokenMetadata(model, L"menu.text");
    ok = Require(metadata.group == L"menu" && metadata.contrastKnown,
                 L"Theme token metadata should expose group, usage, and measured contrast status.") &&
         ok;

    const std::array themeCases = {
        RedConfigure::Workflow::ThemeMassRequest{.recipe = RedConfigure::Workflow::ThemeRecipe::DarkVariant, .keys = {L"menu.background"}},
        RedConfigure::Workflow::ThemeMassRequest{.recipe = RedConfigure::Workflow::ThemeRecipe::LightVariant, .keys = {L"menu.background"}},
        RedConfigure::Workflow::ThemeMassRequest{.recipe = RedConfigure::Workflow::ThemeRecipe::AccentRecolor, .keys = {L"menu.background"}},
        RedConfigure::Workflow::ThemeMassRequest{.recipe = RedConfigure::Workflow::ThemeRecipe::SoftenedSelections, .keys = {L"menu.background"}},
        RedConfigure::Workflow::ThemeMassRequest{.recipe = RedConfigure::Workflow::ThemeRecipe::IncreasedContrast, .keys = {L"menu.background"}},
        RedConfigure::Workflow::ThemeMassRequest{.recipe = RedConfigure::Workflow::ThemeRecipe::SemanticStatusColors,
                                                 .keys = {L"folderView.warningBackground"}},
        RedConfigure::Workflow::ThemeMassRequest{.recipe = RedConfigure::Workflow::ThemeRecipe::SetAlpha,
                                                 .keys = {L"menu.background"},
                                                 .alphaPercent = 37u},
        RedConfigure::Workflow::ThemeMassRequest{.recipe = RedConfigure::Workflow::ThemeRecipe::ReplaceReference,
                                                 .keys = {L"menu.text"},
                                                 .argument = L"palette.accent=palette.warning"},
        RedConfigure::Workflow::ThemeMassRequest{.recipe = RedConfigure::Workflow::ThemeRecipe::ConvertSolidsToReferences,
                                                 .keys = {L"menu.background"},
                                                 .argument = L"accent"},
        RedConfigure::Workflow::ThemeMassRequest{.recipe = RedConfigure::Workflow::ThemeRecipe::RemoveOverrides,
                                                 .keys = {L"menu.background"}},
    };
    for (const RedConfigure::Workflow::ThemeMassRequest& request : themeCases)
    {
        RedConfigure::Themes::ThemePreviewModel candidate;
        candidate.SetTheme(theme);
        const auto preview = RedConfigure::Workflow::PreviewThemeMassChange(candidate, request);
        ok = Require(preview.result == RedConfigure::Workflow::BatchApprovalResult::Ready && ! preview.changes.empty(),
                     L"Every theme mass recipe should produce a validated ready preview.") &&
             ok;
        ok = Require(RedConfigure::Workflow::ApplyThemeMassChange(candidate, preview) == RedConfigure::Workflow::BatchApprovalResult::Applied,
                     L"Every ready theme mass recipe should apply atomically.") &&
             ok;
    }

    const auto partialReference = RedConfigure::Workflow::PreviewThemeMassChange(
        model,
        {.recipe = RedConfigure::Workflow::ThemeRecipe::ReplaceReference,
         .keys = {L"menu.text"},
         .argument = L"palette.acc=palette.warning"});
    ok = Require(partialReference.result == RedConfigure::Workflow::BatchApprovalResult::NoChanges,
                 L"Reference replacement should not rewrite partial reference-name matches.") &&
         ok;
    const auto nonDirectConversion = RedConfigure::Workflow::PreviewThemeMassChange(
        model,
        {.recipe = RedConfigure::Workflow::ThemeRecipe::ConvertSolidsToReferences, .keys = {L"menu.text"}, .argument = L"accent"});
    ok = Require(nonDirectConversion.result == RedConfigure::Workflow::BatchApprovalResult::NoChanges,
                 L"Solid conversion should preserve non-direct theme sources.") &&
         ok;
    const auto missingPalette = RedConfigure::Workflow::PreviewThemeMassChange(
        model,
        {.recipe = RedConfigure::Workflow::ThemeRecipe::ConvertSolidsToReferences, .keys = {L"menu.background"}, .argument = L"missing"});
    ok = Require(missingPalette.result == RedConfigure::Workflow::BatchApprovalResult::Invalid,
                 L"Solid conversion should reject a missing palette target before preview.") &&
         ok;
    const auto malformedReference = RedConfigure::Workflow::PreviewThemeMassChange(
        model,
        {.recipe = RedConfigure::Workflow::ThemeRecipe::ReplaceReference, .keys = {L"menu.text"}, .argument = L"palette.accent"});
    ok = Require(malformedReference.result == RedConfigure::Workflow::BatchApprovalResult::Invalid,
                 L"Reference replacement should reject malformed arguments before preview.") &&
         ok;
    const auto stalePreview = RedConfigure::Workflow::PreviewThemeMassChange(
        model,
        {.recipe = RedConfigure::Workflow::ThemeRecipe::SetAlpha, .keys = {L"menu.background"}, .alphaPercent = 50u});
    ok = Require(model.TryEditOverride(L"menu.background", L"#112233"), L"Theme stale-preview fixture should accept an intervening edit.") && ok;
    ok = Require(RedConfigure::Workflow::ApplyThemeMassChange(model, stalePreview) == RedConfigure::Workflow::BatchApprovalResult::Stale &&
                     model.GetAuthoredColorText(L"menu.background") == L"#112233",
                 L"A stale theme preview should reject atomically without overwriting the newer value.") &&
         ok;
    const std::wstring maximumId = std::wstring(L"user/") + std::wstring(59u, L'a');
    const std::wstring maximumName(64u, L'N');
    const auto boundaryDuplicate = RedConfigure::Workflow::BuildDuplicateThemeCandidate(maximumId, maximumName, L"Copy", 100u);
    ok = Require(boundaryDuplicate.has_value() && boundaryDuplicate->id.size() == 64u && boundaryDuplicate->name.size() == 64u,
                 L"Duplicate theme candidate generation should reserve suffix space at both 64-character boundaries.") &&
         ok;
    return ok;
}

[[nodiscard]] bool TestRedConfigureUiPageCreationAndSwitching()
{
    RedConfigure::RedConfigureSession session;
    RedConfigure::Ui::RedConfigureRootCreateResult result =
        RedConfigure::Ui::CreateRedConfigureRoot(GetModuleHandleW(nullptr), session, std::filesystem::path(L"C:\\RedConfigureUiSmoke"));
    bool ok = Require(result.control != nullptr && result.controller != nullptr, L"RedConfigure UI smoke should create the workbench root and controller.");
    if (! ok)
    {
        return false;
    }
    result.control->SetBounds(D2D1::RectF(0.0f, 0.0f, 1440.0f, 900.0f));
    for (size_t pageIndex = 0u; pageIndex < RedConfigure::GetPageDefinitions().size(); ++pageIndex)
    {
        result.controller->SelectPageForTest(pageIndex);
    }
    result.controller->SelectPageForTest(0u);
    return true;
}

[[nodiscard]] bool TestThemePreviewModelPaletteOperationsAndDynamicInspection()
{
    Common::Settings::ThemeDefinition theme;
    theme.id = L"user/palette-operations";
    theme.name = L"Palette Operations";
    theme.baseThemeId = L"builtin/dark";

    const auto addSource = [](auto& destination, std::wstring key, std::wstring_view text)
    {
        Common::Settings::ThemeColorSource source;
        if (FAILED(Common::Settings::ParseThemeColorSource(text, source))) return false;
        destination.emplace(std::move(key), std::move(source));
        return true;
    };

    bool ok = true;
    ok = Require(addSource(theme.palette, L"accent", L"#336699"), L"Palette fixture accent should parse.") && ok;
    ok = Require(addSource(theme.palette, L"alternate", L"#CC8844"), L"Palette fixture alternate should parse.") && ok;
    ok = Require(addSource(theme.colors, L"app.accent", L"ref(palette.accent)"), L"Palette fixture semantic reference should parse.") && ok;
    ok = Require(addSource(theme.colors,
                           L"folderView.itemBackgroundSelected",
                           L"seededChoice(runtime.seed,palette.accent,palette.alternate)"),
                 L"Palette fixture dynamic source should parse.") && ok;

    RedConfigure::Themes::ThemePreviewModel model;
    model.SetTheme(theme);
    ok = Require(model.GetEffectiveColor(L"palette.accent").value_or(0u) == 0xFF336699u,
                 L"Palette entries should expose their effective swatch.") && ok;
    ok = Require(model.GetDependencies(L"app.accent") == std::vector<std::wstring>{L"palette.accent"},
                 L"Semantic dependency inspection should expose palette references.") && ok;
    ok = Require(model.GetAffected(L"palette.accent").size() == 2u,
                 L"Reverse dependency inspection should expose every dependent source.") && ok;
    ok = Require(model.GetEvaluationPhase(L"folderView.itemBackgroundSelected") == Common::Settings::ThemeColorEvaluationPhase::Paint,
                 L"Dynamic selection sources should report paint-time evaluation.") && ok;

    model.SetPreviewSeed(1u);
    ok = Require(model.GetEffectiveColor(L"folderView.itemBackgroundSelected").value_or(0u) == 0xFFCC8844u,
                 L"The fixed preview seed should evaluate the same candidate order as runtime.") && ok;
    ok = Require(model.RenamePaletteEntry(L"accent", L"brand"), L"Palette rename should rewrite all dependents atomically.") && ok;
    ok = Require(model.GetAuthoredColorText(L"app.accent") == L"ref(palette.brand)",
                 L"Palette rename should preserve authored expressions with the new reference.") && ok;
    ok = Require(! model.ResetOverride(L"palette.brand") && ! model.GetLastError().empty(),
                 L"Deleting a referenced palette entry should be blocked with a diagnostic.") && ok;
    ok = Require(model.TryEditOverride(L"app.accent", L"#112233"), L"A dependent semantic token should remain directly editable.") && ok;
    ok = Require(model.TryEditOverride(L"folderView.itemBackgroundSelected", L"#445566"),
                 L"A dynamic dependent should be replaceable by a static source.") && ok;
    ok = Require(model.ResetOverride(L"palette.brand"), L"An unreferenced palette entry should be removable.") && ok;
    ok = Require(model.TryEditOverride(L"menu.background", L"#445566"), L"Repeated-literal fixture should be editable.") && ok;
    ok = Require(model.CreatePaletteEntry(L"sharedSelection", L"#445566", true),
                 L"Creating a palette entry should optionally convert every matching direct source.") && ok;
    ok = Require(model.GetAuthoredColorText(L"menu.background") == L"ref(palette.sharedSelection)" &&
                     model.GetAuthoredColorText(L"folderView.itemBackgroundSelected") == L"ref(palette.sharedSelection)" &&
                     model.GetEffectiveColor(L"menu.background").value_or(0u) == 0xFF445566u,
                 L"Repeated-literal conversion should preserve effective colors while removing authored redundancy.") && ok;
    ok = Require(model.WrapSourceWithTransform(L"menu.background", RedConfigure::Themes::ThemeSourceTransform::Darken10),
                 L"Batch-style transforms should preserve the original source in a generated palette recipe.") && ok;
    ok = Require(model.GetAuthoredColorText(L"menu.background").starts_with(L"darken(palette.source_menu_background") &&
                     model.GetEffectiveColor(L"menu.background").value_or(0u) == 0xFF3D4D5Cu,
                 L"A darken transform should remain authored as a function instead of flattening to hex.") && ok;

    Common::Settings::ThemeDefinition fanoutTheme;
    fanoutTheme.id = L"user/preview-fanout";
    fanoutTheme.name = L"Preview Fanout";
    fanoutTheme.baseThemeId = L"builtin/dark";
    fanoutTheme.palette.emplace(L"root", Common::Settings::ThemeColorSource(0xFF102030u));
    for (size_t index = 0u; index < 512u; ++index)
    {
        Common::Settings::ThemeColorSource source;
        source.kind = Common::Settings::ThemeColorSourceKind::Reference;
        source.references.push_back(L"palette.root");
        fanoutTheme.colors.emplace(std::format(L"selftest.preview.{}", index), std::move(source));
    }
    RedConfigure::Themes::ThemePreviewModel fanoutModel;
    fanoutModel.SetTheme(fanoutTheme);
    ok = Require(fanoutModel.GetAffected(L"palette.root").size() == 512u &&
                     fanoutModel.TryEditOverride(L"palette.root", L"#405060") &&
                     fanoutModel.GetEffectiveColor(L"selftest.preview.511").value_or(0u) == 0xFF405060u,
                 L"A worst-allowed 512-token fan-out edit should recompute once and update every dependent preview color.") && ok;
    std::vector<std::wstring> fanoutKeys;
    fanoutKeys.reserve(512u);
    for (size_t index = 0u; index < 512u; ++index)
    {
        fanoutKeys.push_back(std::format(L"selftest.preview.{}", index));
    }
    const auto massPreviewStarted = std::chrono::steady_clock::now();
    const auto massPreview = RedConfigure::Workflow::PreviewThemeMassChange(
        fanoutModel,
        {.recipe = RedConfigure::Workflow::ThemeRecipe::SetAlpha, .keys = std::move(fanoutKeys), .alphaPercent = 80u});
    const auto massPreviewElapsed = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - massPreviewStarted);
    std::wcout << L"redconfigure.theme.mass_preview keys=512 elapsedMicros=" << massPreviewElapsed.count() << L'\n';
    ok = Require(massPreview.result == RedConfigure::Workflow::BatchApprovalResult::Ready && massPreview.changes.size() == 512u,
                 L"The worst-allowed 512-token theme mass operation should validate a complete preview.") &&
         ok;
    ok = Require(massPreviewElapsed < std::chrono::milliseconds(100),
                 L"The worst-allowed 512-token theme mass preview should remain below the 100 ms explicit-action budget.") &&
         ok;
    return ok;
}

[[nodiscard]] bool TestThemeColorSuggestionsGuideExpressionEditing()
{
    const std::vector<std::wstring> suggestions =
        RedConfigure::BuildThemeColorSuggestions(L"folderView.itemBackgroundSelected", L"menu.background", 0xFF2ECC71u);

    const auto contains = [&suggestions](std::wstring_view expected) noexcept
    { return std::find(suggestions.begin(), suggestions.end(), expected) != suggestions.end(); };

    bool ok = true;
    ok      = Require(contains(L"#2ECC71"), L"Theme color suggestions should include the current direct color.") && ok;
    ok      = Require(contains(L"ref(app.accent)"), L"Theme color suggestions should include an accent reference.") && ok;
    ok      = Require(contains(L"darken(app.accent,20%)"), L"Theme color suggestions should include a darken expression template.") && ok;
    ok      = Require(contains(L"blend(menu.background,app.accent,16%)"),
                      L"Theme color suggestions should include a blend expression using the previously selected color.") &&
              ok;
    ok      = Require(contains(L"ref(menu.background)"), L"Theme color suggestions should include the previous color reference.") && ok;
    ok      = Require(contains(L"perceptualTone(app.accent,60)") && contains(L"ensureContrast(menu.text,menu.background,4.5)") &&
                          contains(L"harmonize(app.accent,navigation.accent,25%)") && contains(L"systemColor(accent)"),
                      L"Theme color suggestions should expose the approved load-time and event-time function templates.") && ok;
    ok      = Require(contains(L"seededRainbow(runtime.seed,85%,75%,100%,0)") &&
                          contains(L"seededChoice(runtime.seed,app.accent,navigation.accent)"),
                      L"The allowlisted selection token should expose deterministic paint-time function templates.") && ok;
    return ok;
}

[[nodiscard]] bool TestThemeColorKeyFilterNarrowsKeysCaseInsensitively()
{
    const std::vector<std::wstring> keys = {
        L"app.accent",
        L"menu.background",
        L"menu.selectionBg",
        L"folderView.background",
        L"viewer.diff.addedBackground",
    };

    bool ok                            = true;
    std::vector<std::wstring> filtered = RedConfigure::FilterThemeColorKeys(keys, L"MENU");
    ok                                 = Require(filtered == std::vector<std::wstring>{L"menu.background", L"menu.selectionBg"},
                                                 L"Theme color key filter should match key groups case-insensitively.") &&
                                         ok;

    filtered = RedConfigure::FilterThemeColorKeys(keys, L"background");
    ok       = Require(filtered == std::vector<std::wstring>{L"menu.background", L"folderView.background", L"viewer.diff.addedBackground"},
                       L"Theme color key filter should match substrings anywhere in the key.") &&
               ok;

    filtered = RedConfigure::FilterThemeColorKeys(keys, L"missing");
    ok       = Require(filtered.empty(), L"Theme color key filter should produce an empty list when no keys match.") && ok;
    return ok;
}

[[nodiscard]] bool TestThemePreviewHitSelectionUsesSmallestRegionAndCycles()
{
    const std::vector<RedConfigure::ThemePreviewHitCandidate> candidates = {
        RedConfigure::ThemePreviewHitCandidate{.key = L"navigation.background", .left = 0.0f, .top = 0.0f, .right = 200.0f, .bottom = 200.0f},
        RedConfigure::ThemePreviewHitCandidate{.key = L"app.accent", .left = 0.0f, .top = 0.0f, .right = 6.0f, .bottom = 200.0f},
        RedConfigure::ThemePreviewHitCandidate{.key = L"menu.selectionBg", .left = 60.0f, .top = 10.0f, .right = 120.0f, .bottom = 40.0f},
        RedConfigure::ThemePreviewHitCandidate{.key = L"menu.background", .left = 40.0f, .top = 0.0f, .right = 240.0f, .bottom = 60.0f},
    };

    bool ok = true;
    ok      = Require(RedConfigure::SelectThemePreviewHitKey(candidates, 3.0f, 50.0f, {}) == L"app.accent",
                      L"Theme preview hit selection should prefer the smallest visible region.") &&
              ok;
    ok      = Require(RedConfigure::SelectThemePreviewHitKey(candidates, 3.0f, 50.0f, L"app.accent") == L"navigation.background",
                      L"Theme preview repeated clicks should cycle to the containing region.") &&
              ok;
    ok      = Require(RedConfigure::SelectThemePreviewHitKey(candidates, 80.0f, 20.0f, {}) == L"menu.selectionBg",
                      L"Theme preview hit selection should prefer nested menu selection over menu background.") &&
              ok;
    ok      = Require(RedConfigure::SelectThemePreviewHitKey(candidates, 500.0f, 500.0f, {}).empty(),
                      L"Theme preview hit selection should return no key outside all regions.") &&
              ok;
    return ok;
}

[[nodiscard]] bool TestRedConfigureSessionExportsFirstUsableFiles()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureSessionTest", ec);
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
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"App.rc",
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
    ok      = Require(WriteTestTextFile(tempRoot / L"App" / L"Lang" / L"fr-FR" / L"App-fr-FR.rc",
                                        "#include \"resource.h\"\nSTRINGTABLE\nBEGIN\n    IDS_HELLO \"Bonjour {0}\"\nEND\n"),
                      L"Failed to write session target rc fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Zeta" / L"Zeta.vcxproj",
                                        R"xml(<?xml version="1.0" encoding="utf-8"?>
<Project>
  <ItemGroup>
    <ResourceCompile Include="Zeta.rc" />
  </ItemGroup>
</Project>)xml"),
                      L"Failed to write second session project fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Zeta" / L"Zeta.rc", "STRINGTABLE\nBEGIN\n    IDS_ZETA \"Zeta\"\nEND\n"),
                      L"Failed to write second source rc fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Themes" / L"Session.theme.json5",
                                        R"json5({
  formatVersion: 2,
  id: "user/session",
  name: "Session",
  baseThemeId: "builtin/dark",
  colors: {
    "app.accent": "#336699",
  },
})json5"),
                      L"Failed to write session theme fixture.") &&
              ok;
    ok      = Require(WriteTestTextFile(tempRoot / L"Themes" / L"Zed.theme.json5",
                                        R"json5({
  formatVersion: 2,
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
        ok                = Require(first.id == L"IDS_HELLO", L"Session should preserve string ids.") && ok;
        ok                = Require(first.targetText == L"Bonjour {0}", L"Session should merge target translations.") && ok;
    }
    ok = Require(SUCCEEDED(session.SetActiveResourceOwner(1u)), L"Session should switch active resource owners.") && ok;
    ok = Require(! session.GetTranslations().empty() && session.GetTranslations().front().id == L"IDS_ZETA",
                 L"Session should reload translations when the owner changes.") &&
         ok;
    ok = Require(session.SetActiveTheme(1u), L"Session should switch active themes.") && ok;
    ok = Require(session.GetThemePreviewModel().GetTheme().id == L"user/zed", L"Session should update the preview theme when switched.") && ok;
    ok = Require(session.DuplicateActiveTheme(L"invalid", L"Copy") == RedConfigure::DuplicateThemeResult::InvalidId,
                 L"Duplicate theme creation should distinguish an invalid identifier.") &&
         ok;
    ok = Require(session.DuplicateActiveTheme(L"user/new", L"") == RedConfigure::DuplicateThemeResult::InvalidName,
                 L"Duplicate theme creation should distinguish an invalid name.") &&
         ok;
    ok = Require(session.DuplicateActiveTheme(L"user/zed", L"Copy") == RedConfigure::DuplicateThemeResult::Collision,
                 L"Duplicate theme creation should distinguish an identifier collision.") &&
         ok;
    ok = Require(session.DuplicateActiveTheme(L"user/zed-copy", L"Zed Copy") == RedConfigure::DuplicateThemeResult::Created &&
                     ! session.GetThemeCatalog().themes.empty() &&
                     session.GetThemeCatalog().themes.back().origin == RedConfigure::Themes::ThemeCatalogOrigin::User && session.Undo(),
                 L"A successful duplicate theme should be created and undone as one session edit.") &&
         ok;
    ok = Require(SUCCEEDED(session.SetActiveResourceOwner(0u)), L"Session should switch back to the first resource owner.") && ok;

    ok = Require(session.UpdateTranslation(0u, L"Salut {0}"), L"Session should accept a valid translation edit.") && ok;
    ok = Require(! session.UpdateThemeColor(L"app.accent", L"not-a-color"), L"Session should reject invalid theme colors.") && ok;
    ok = Require(session.UpdateThemeColor(L"app.accent", L"#123456"), L"Session should accept valid theme colors.") && ok;
    const auto themeBatch = RedConfigure::Workflow::PreviewThemeMassChange(
        session.GetThemePreviewModel(),
        {.recipe = RedConfigure::Workflow::ThemeRecipe::SetAlpha, .keys = {L"app.accent"}, .alphaPercent = 60u});
    ok = Require(session.ApplyThemeMassChange(themeBatch) == RedConfigure::Workflow::BatchApprovalResult::Applied,
                 L"A validated theme batch should apply through the session.") &&
         ok;
    ok = Require(session.Undo() && session.GetThemePreviewModel().GetAuthoredColorText(L"app.accent") == L"#123456",
                 L"A successful theme batch should be exactly one Undo step.") &&
         ok;
    RedConfigure::Ui::ThemesPagePresenter themesPresenter;
    const RedConfigure::Workflow::ThemeMassRequest presenterThemeRequest{
        .recipe = RedConfigure::Workflow::ThemeRecipe::SetAlpha, .keys = {L"app.accent"}, .alphaPercent = 50u};
    const auto firstThemePresentation = themesPresenter.Execute(session, presenterThemeRequest);
    RedConfigure::Workflow::ThemeMassRequest changedThemeRequest = presenterThemeRequest;
    changedThemeRequest.alphaPercent = 60u;
    const auto changedThemePresentation = themesPresenter.Execute(session, changedThemeRequest);
    ok = Require(firstThemePresentation.phase == RedConfigure::Ui::BatchInteractionPhase::Preview &&
                     changedThemePresentation.phase == RedConfigure::Ui::BatchInteractionPhase::Preview,
                 L"Changing a theme argument or alpha value should replace approval with a new preview instead of applying.") &&
         ok;
    const auto presenterThemeApply = themesPresenter.Execute(session, changedThemeRequest);
    ok = Require(presenterThemeApply.phase == RedConfigure::Ui::BatchInteractionPhase::Apply &&
                     presenterThemeApply.result == RedConfigure::Workflow::BatchApprovalResult::Applied && session.Undo(),
                 L"An unchanged theme presenter request should apply once and remain one Undo step.") &&
         ok;
    ok = Require(session.UpdateThemeColor(L"folderView.itemBackgroundSelected", L"darken(app.accent,50%)"),
                 L"Session should accept theme expression color edits.") &&
         ok;

    std::wstring rcPreview;
    std::string themePreview;
    ok = Require(SUCCEEDED(session.BuildLocalizationExportText(rcPreview)), L"Session should build a localization export preview.") && ok;
    ok = Require(SUCCEEDED(session.BuildThemeExportText(themePreview)), L"Session should build a theme export preview.") && ok;
    ok = Require(rcPreview.find(L"Salut {0}") != std::wstring::npos, L"Localization export preview should contain edited translations.") && ok;
    ok = Require(themePreview.find("\"app.accent\": \"#123456\"") != std::string::npos, L"Theme export preview should contain edited colors.") && ok;
    ok = Require(themePreview.find("\"folderView.itemBackgroundSelected\": \"darken(app.accent,0.5)\"") != std::string::npos,
                 L"Theme export preview should preserve expression-authored colors.") &&
         ok;

    const std::filesystem::path rcPath    = tempRoot / L"Out" / L"App-fr-FR.rc";
    const std::filesystem::path themePath = tempRoot / L"Out" / L"Session.theme.json5";
    ok                                    = Require(SUCCEEDED(session.ExportLocalization(rcPath)), L"Session should export a satellite rc file.") && ok;
    ok                                    = Require(SUCCEEDED(session.ExportTheme(themePath)), L"Session should export a theme json5 file.") && ok;
    ok                                    = Require(std::filesystem::exists(rcPath, ec), L"Exported rc file should exist.") && ok;
    ok                                    = Require(std::filesystem::exists(themePath, ec), L"Exported theme file should exist.") && ok;

    std::filesystem::remove_all(tempRoot, ec);
    return ok;
}

[[nodiscard]] bool TestRedConfigureSessionReadsBomlessUtf16LeRcFiles()
{
    std::error_code ec;
    const std::filesystem::path tempRoot = AcquireRedConfigureTestSandbox(L"RedConfigureBomlessUtf16LeSessionTest", ec);
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
    ok      = Require(WriteTestUtf16LeTextFile(tempRoot / L"App" / L"App.rc",
                                               LR"rc(#include "resource.h"
STRINGTABLE
BEGIN
    IDS_HELLO "Héllo {0}"
END
)rc",
                                               false),
                      L"Failed to write BOM-less UTF-16 LE source rc fixture.") &&
              ok;
    ok      = Require(WriteTestUtf16LeTextFile(tempRoot / L"App" / L"Lang" / L"fr-FR" / L"App-fr-FR.rc",
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
        ok = Require(session.GetTranslations().front().sourceText == L"Héllo {0}", L"Session should preserve non-ASCII source text from BOM-less UTF-16 LE.") &&
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
    const std::wstring perfJsonlPath = GetEnvironmentString(L"REDCONFIGURE_PERF_JSONL");
    if (! perfJsonlPath.empty())
    {
        Debug::Perf::ConfigureJsonlOutput(perfJsonlPath,
                                          L"RedConfigure repo-sized scan and validation",
                                          L"Debug x64",
                                          {},
                                          {},
                                          GetEnvironmentString(L"REDCONFIGURE_PERF_MACHINE_HASH"),
                                          GetEnvironmentString(L"REDCONFIGURE_PERF_RUN_ID"));
    }
    auto clearPerfOutput = wil::scope_exit([] { Debug::Perf::ClearJsonlOutput(); });
    if (! TestPageDefinitions())
    {
        return 1;
    }
    if (! TestResolveWorkspaceRootForLaunchPath())
    {
        return 1;
    }
    if (! TestBoundedBinaryFileReader())
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
    if (! TestThemeExpressionLanguageAndRuntimeSources())
    {
        return 1;
    }
    if (! TestThemeExpressionNumbersAreLocaleInvariant())
    {
        return 1;
    }
    if (! TestSettingsStoreThemeDirectoryUsesSharedParser())
    {
        return 1;
    }
    if (! TestShippedThemesUseVersion2Sources())
    {
        return 1;
    }
    if (! TestSettingsStoreInlineThemesUseSharedParser())
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
    if (! TestRedConfigureSessionBuildsAllOwnerAllCultureLocalizationReview())
    {
        return 1;
    }
    if (! TestLocalizationReviewViewFiltersSearchesAndSorts())
    {
        return 1;
    }
    if (! TestLocalizationReviewEditingAndExportPreviews())
    {
        return 1;
    }
    if (! TestLocalizationReviewSkipsBadResourceFilesAndKeepsWorkspaceOpen())
    {
        return 1;
    }
    if (! TestLocalizationReviewGridModelShowsOwnersEnglishAndVisibleCultures())
    {
        return 1;
    }
    if (! TestLocalizationReviewSurfacesMissingTranslations())
    {
        return 1;
    }
    if (! TestLocalizationReviewCanCreateAndExportNewCulture())
    {
        return 1;
    }
    if (! TestLocalizationReviewProjectionPerformanceMetric())
    {
        return 1;
    }
    if (! TestRedConfigureRepoSizedScanAndValidationPerformance())
    {
        return 1;
    }
    if (! TestSplashScreenCloseGuardSuppressesPendingOpen())
    {
        return 1;
    }
    if (! TestRcWriterAndMerge())
    {
        return 1;
    }
    if (! TestGeneratedRcCompilesWhenWindowsSdkIsAvailable())
    {
        return 1;
    }
    if (! TestRedConfigureWorkflowModels())
    {
        return 1;
    }
    if (! TestRedConfigureUiPageCreationAndSwitching())
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
    if (! TestThemePreviewModelPaletteOperationsAndDynamicInspection())
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
