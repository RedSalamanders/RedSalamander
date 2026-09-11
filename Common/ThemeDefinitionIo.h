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
    MissingOrInvalidFormatVersion,
    UnsupportedFormatVersion,
    MissingOrInvalidId,
    InvalidId,
    MissingOrInvalidName,
    MissingOrInvalidBaseThemeId,
    InvalidBaseThemeId,
    PaletteNotObject,
    TooManyPaletteEntries,
    InvalidPaletteName,
    DuplicatePaletteName,
    ColorsMissingOrNotObject,
    TooManyColorEntries,
    InvalidColorKey,
    DuplicateColorKey,
    ColorValueNotString,
    InvalidColorValue,
    OutOfMemory,
};

enum class ThemeDefinitionParseMode : uint8_t
{
    StrictFile,
    LenientInline,
};

COMMON_API bool IsValidThemeColorKey(std::wstring_view key) noexcept;
COMMON_API bool IsValidThemePaletteName(std::wstring_view name) noexcept;
COMMON_API bool IsValidUserThemeId(std::wstring_view id) noexcept;
COMMON_API bool IsBuiltinThemeId(std::wstring_view themeId) noexcept;

COMMON_API HRESULT ParseThemeDefinitionJson5(std::string_view jsonText,
                                             ThemeDefinition& outTheme,
                                             ThemeDefinitionIoError* outError,
                                             std::wstring* outMessage) noexcept;

COMMON_API HRESULT ParseThemeDefinitionFromValue(const JsonValue& value,
                                                  ThemeDefinition& outTheme,
                                                  ThemeDefinitionParseMode mode,
                                                  ThemeDefinitionIoError* outError,
                                                  std::wstring* outMessage,
                                                  uint32_t* outSkippedColorEntries = nullptr) noexcept;

COMMON_API HRESULT BuildThemeDefinitionJson5(const ThemeDefinition& theme, std::string& outJson) noexcept;
} // namespace Common::Settings
