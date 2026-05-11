#pragma once

#include "SettingsStore.h"

#include <cstdint>
#include <string>
#include <string_view>

namespace Common::Settings
{
enum class ThemeDefinitionIoError : uint8_t
{
    None,
    EmptyInput,
    ParseFailed,
    RootNotObject,
    MissingOrInvalidId,
    InvalidId,
    MissingOrInvalidName,
    MissingOrInvalidBaseThemeId,
    InvalidBaseThemeId,
    ColorsMissingOrNotObject,
    InvalidColorKey,
    ColorValueNotString,
    InvalidColorValue,
    OutOfMemory,
};

COMMON_API bool IsValidThemeColorKey(std::wstring_view key) noexcept;
COMMON_API bool IsValidUserThemeId(std::wstring_view id) noexcept;
COMMON_API bool IsBuiltinThemeId(std::wstring_view themeId) noexcept;

COMMON_API HRESULT ParseThemeDefinitionJson5(std::string_view jsonText,
                                             ThemeDefinition& outTheme,
                                             ThemeDefinitionIoError* outError,
                                             std::wstring* outMessage) noexcept;

COMMON_API HRESULT BuildThemeDefinitionJson5(const ThemeDefinition& theme, std::string& outJson) noexcept;
} // namespace Common::Settings
