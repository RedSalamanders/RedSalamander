#include "ThemeDefinitionIo.h"
#include "Helpers.h"

#include <algorithm>
#include <array>
#include <cstdlib>
#include <format>
#include <limits>
#include <string>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026, C5027
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#pragma warning(push)
// yyjson currently trips analyzer null-deref warnings in iterator internals.
#pragma warning(disable : 28182)
#include <yyjson.h>
#pragma warning(pop)

namespace
{
constexpr std::array<std::wstring_view, 5> kBuiltinThemeIds = {{
    L"builtin/system",
    L"builtin/light",
    L"builtin/dark",
    L"builtin/rainbow",
    L"builtin/highContrast",
}};

[[nodiscard]] HRESULT InvalidData(Common::Settings::ThemeDefinitionIoError error,
                                  Common::Settings::ThemeDefinitionIoError* outError,
                                  std::wstring* outMessage,
                                  std::wstring message = {}) noexcept
{
    if (outError)
    {
        *outError = error;
    }
    if (outMessage)
    {
        *outMessage = std::move(message);
    }
    return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

[[nodiscard]] HRESULT OutOfMemory(Common::Settings::ThemeDefinitionIoError* outError, std::wstring* outMessage) noexcept
{
    if (outError)
    {
        *outError = Common::Settings::ThemeDefinitionIoError::OutOfMemory;
    }
    if (outMessage)
    {
        outMessage->clear();
    }
    return E_OUTOFMEMORY;
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    return Common::Strings::Utf8FromUtf16StrictOrEmpty(text);
}

[[nodiscard]] HRESULT AddStringMember(yyjson_mut_doc* doc, yyjson_mut_val* obj, const char* key, std::string_view value) noexcept
{
    yyjson_mut_val* valueNode = yyjson_mut_strncpy(doc, value.data(), value.size());
    if (! valueNode)
    {
        return E_OUTOFMEMORY;
    }
    return yyjson_mut_obj_add_val(doc, obj, key, valueNode) ? S_OK : E_OUTOFMEMORY;
}
} // namespace

namespace Common::Settings
{
bool IsValidThemeColorKey(std::wstring_view key) noexcept
{
    if (key.empty() || key.size() > 64)
    {
        return false;
    }

    for (const wchar_t ch : key)
    {
        const bool ok = (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z') || (ch >= L'0' && ch <= L'9') || ch == L'_' || ch == L'.' || ch == L'-';
        if (! ok)
        {
            return false;
        }
    }

    return true;
}

bool IsValidThemePaletteName(std::wstring_view name) noexcept
{
    return IsValidThemeColorKey(name) && name.find(L'.') == std::wstring_view::npos;
}

bool IsValidUserThemeId(std::wstring_view id) noexcept
{
    constexpr std::wstring_view prefix = L"user/";
    if (id.rfind(prefix, 0) != 0)
    {
        return false;
    }

    const std::wstring_view suffix = id.substr(prefix.size());
    if (suffix.empty() || suffix.size() > 64)
    {
        return false;
    }

    const wchar_t first = suffix.front();
    const bool firstOk  = (first >= L'A' && first <= L'Z') || (first >= L'a' && first <= L'z') || (first >= L'0' && first <= L'9');
    if (! firstOk)
    {
        return false;
    }

    for (const wchar_t ch : suffix)
    {
        const bool ok = (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z') || (ch >= L'0' && ch <= L'9') || ch == L'_' || ch == L'.' || ch == L'-';
        if (! ok)
        {
            return false;
        }
    }

    return true;
}

bool IsBuiltinThemeId(std::wstring_view themeId) noexcept
{
    return std::find(kBuiltinThemeIds.begin(), kBuiltinThemeIds.end(), themeId) != kBuiltinThemeIds.end();
}

namespace
{
[[nodiscard]] const JsonValue* FindJsonMember(const JsonObject& object, std::string_view key) noexcept
{
    const auto found = std::find_if(object.members.begin(), object.members.end(), [&](const auto& member) noexcept { return member.first == key; });
    return found == object.members.end() ? nullptr : &found->second;
}

[[nodiscard]] bool ReadJsonString(const JsonObject& object, std::string_view key, std::wstring& out) noexcept
{
    const JsonValue* value = FindJsonMember(object, key);
    if (! value)
    {
        return false;
    }
    const auto* text = std::get_if<std::string>(&value->value);
    if (! text || text->empty())
    {
        return false;
    }
    out = Utf16FromUtf8(*text);
    return ! out.empty();
}

[[nodiscard]] bool ReadJsonUInt32(const JsonObject& object, std::string_view key, uint32_t& out) noexcept
{
    const JsonValue* value = FindJsonMember(object, key);
    if (! value)
    {
        return false;
    }
    if (const auto* unsignedValue = std::get_if<uint64_t>(&value->value); unsignedValue && *unsignedValue <= std::numeric_limits<uint32_t>::max())
    {
        out = static_cast<uint32_t>(*unsignedValue);
        return true;
    }
    if (const auto* signedValue = std::get_if<int64_t>(&value->value);
        signedValue && *signedValue >= 0 && static_cast<uint64_t>(*signedValue) <= std::numeric_limits<uint32_t>::max())
    {
        out = static_cast<uint32_t>(*signedValue);
        return true;
    }
    return false;
}

[[nodiscard]] HRESULT ParseSourceObject(const JsonObject& object,
                                        bool palette,
                                        ThemeDefinitionParseMode mode,
                                        std::unordered_map<std::wstring, ThemeColorSource>& out,
                                        ThemeDefinitionIoError* outError,
                                        std::wstring* outMessage,
                                        uint32_t& skipped) noexcept
{
    for (const auto& [keyUtf8, value] : object.members)
    {
        const std::wstring key = Utf16FromUtf8(keyUtf8);
        const bool validKey    = palette ? IsValidThemePaletteName(key) : IsValidThemeColorKey(key);
        const bool duplicate   = out.contains(key);
        const auto* sourceText = std::get_if<std::string>(&value.value);
        ThemeColorSource source;
        std::wstring parseMessage;
        const bool validSource    = sourceText && SUCCEEDED(ParseThemeColorSource(Utf16FromUtf8(*sourceText), source, &parseMessage));
        const bool paletteDynamic = palette && validSource && (IsPaintTimeThemeColorSource(source) || IsEventTimeThemeColorSource(source));
        if (! validKey || duplicate || ! sourceText || ! validSource || paletteDynamic)
        {
            if (mode == ThemeDefinitionParseMode::StrictFile)
            {
                const ThemeDefinitionIoError error =
                    ! validKey     ? (palette ? ThemeDefinitionIoError::InvalidPaletteName : ThemeDefinitionIoError::InvalidColorKey)
                    : duplicate    ? (palette ? ThemeDefinitionIoError::DuplicatePaletteName : ThemeDefinitionIoError::DuplicateColorKey)
                    : ! sourceText ? ThemeDefinitionIoError::ColorValueNotString
                                   : ThemeDefinitionIoError::InvalidColorValue;
                return InvalidData(error, outError, outMessage, std::move(parseMessage));
            }
            ++skipped;
            continue;
        }
        out.emplace(key, std::move(source));
    }
    return S_OK;
}
} // namespace

HRESULT ParseThemeDefinitionFromValue(const JsonValue& value,
                                      ThemeDefinition& outTheme,
                                      ThemeDefinitionParseMode mode,
                                      ThemeDefinitionIoError* outError,
                                      std::wstring* outMessage,
                                      uint32_t* outSkippedColorEntries) noexcept
{
    outTheme = {};
    if (outError)
    {
        *outError = ThemeDefinitionIoError::None;
    }
    if (outMessage)
    {
        outMessage->clear();
    }
    if (outSkippedColorEntries)
    {
        *outSkippedColorEntries = 0u;
    }

    const auto* objectPtr = std::get_if<JsonValue::ObjectPtr>(&value.value);
    if (! objectPtr || ! *objectPtr)
    {
        return InvalidData(ThemeDefinitionIoError::RootNotObject, outError, outMessage);
    }
    const JsonObject& object = **objectPtr;
    ThemeDefinition parsed;
    if (! ReadJsonUInt32(object, "formatVersion", parsed.formatVersion))
    {
        return InvalidData(ThemeDefinitionIoError::MissingOrInvalidFormatVersion, outError, outMessage, L"Theme formatVersion must be the integer 2.");
    }
    if (parsed.formatVersion != 2u)
    {
        return InvalidData(ThemeDefinitionIoError::UnsupportedFormatVersion,
                           outError,
                           outMessage,
                           std::format(L"Unsupported theme formatVersion {}. Only version 2 is accepted.", parsed.formatVersion));
    }
    if (! ReadJsonString(object, "id", parsed.id))
    {
        return InvalidData(ThemeDefinitionIoError::MissingOrInvalidId, outError, outMessage);
    }
    if (! IsValidUserThemeId(parsed.id))
    {
        return InvalidData(ThemeDefinitionIoError::InvalidId, outError, outMessage);
    }
    if (! ReadJsonString(object, "name", parsed.name))
    {
        return InvalidData(ThemeDefinitionIoError::MissingOrInvalidName, outError, outMessage);
    }
    if (parsed.name.size() > 64u)
    {
        if (mode == ThemeDefinitionParseMode::StrictFile)
        {
            return InvalidData(ThemeDefinitionIoError::MissingOrInvalidName, outError, outMessage);
        }
        parsed.name.resize(64u);
    }
    if (! ReadJsonString(object, "baseThemeId", parsed.baseThemeId))
    {
        return InvalidData(ThemeDefinitionIoError::MissingOrInvalidBaseThemeId, outError, outMessage);
    }
    if (mode == ThemeDefinitionParseMode::StrictFile && ! IsBuiltinThemeId(parsed.baseThemeId))
    {
        return InvalidData(ThemeDefinitionIoError::InvalidBaseThemeId, outError, outMessage);
    }

    uint32_t skipped = 0u;
    if (const JsonValue* paletteValue = FindJsonMember(object, "palette"))
    {
        const auto* palettePtr = std::get_if<JsonValue::ObjectPtr>(&paletteValue->value);
        if (! palettePtr || ! *palettePtr)
        {
            return InvalidData(ThemeDefinitionIoError::PaletteNotObject, outError, outMessage);
        }
        if ((*palettePtr)->members.size() > 128u)
        {
            return InvalidData(ThemeDefinitionIoError::TooManyPaletteEntries, outError, outMessage, L"A theme palette cannot contain more than 128 entries.");
        }
        if (const HRESULT hr = ParseSourceObject(**palettePtr, true, mode, parsed.palette, outError, outMessage, skipped); FAILED(hr))
        {
            return hr;
        }
    }

    const JsonValue* colorsValue = FindJsonMember(object, "colors");
    const auto* colorsPtr        = colorsValue ? std::get_if<JsonValue::ObjectPtr>(&colorsValue->value) : nullptr;
    if (! colorsPtr || ! *colorsPtr)
    {
        return InvalidData(ThemeDefinitionIoError::ColorsMissingOrNotObject, outError, outMessage);
    }
    if ((*colorsPtr)->members.size() > 512u)
    {
        return InvalidData(ThemeDefinitionIoError::TooManyColorEntries, outError, outMessage, L"A theme cannot contain more than 512 semantic color entries.");
    }

    if (const HRESULT hr = ParseSourceObject(**colorsPtr, false, mode, parsed.colors, outError, outMessage, skipped); FAILED(hr))
    {
        return hr;
    }

    if (outSkippedColorEntries)
    {
        *outSkippedColorEntries = skipped;
    }
    if (skipped != 0u && outMessage)
    {
        *outMessage = std::format(L"Skipped {} invalid inline theme color entr{}.", skipped, skipped == 1u ? L"y" : L"ies");
    }
    outTheme = std::move(parsed);
    return S_OK;
}

HRESULT ParseThemeDefinitionJson5(std::string_view jsonText, ThemeDefinition& outTheme, ThemeDefinitionIoError* outError, std::wstring* outMessage) noexcept
{
    if (jsonText.empty())
    {
        return InvalidData(ThemeDefinitionIoError::EmptyInput, outError, outMessage);
    }
    JsonValue value;
    const HRESULT parseHr = ParseJsonValue(jsonText, value);
    if (FAILED(parseHr))
    {
        return parseHr == E_OUTOFMEMORY ? OutOfMemory(outError, outMessage) : InvalidData(ThemeDefinitionIoError::ParseFailed, outError, outMessage);
    }
    return ParseThemeDefinitionFromValue(value, outTheme, ThemeDefinitionParseMode::StrictFile, outError, outMessage);
}

HRESULT BuildThemeDefinitionJson5(const ThemeDefinition& theme, std::string& outJson) noexcept
{
    outJson.clear();

    if (theme.formatVersion != 2u || ! IsValidUserThemeId(theme.id) || theme.name.empty() || theme.name.size() > 64u || theme.palette.size() > 128u ||
        theme.colors.size() > 512u || ! IsBuiltinThemeId(theme.baseThemeId))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    wil::unique_any<yyjson_mut_doc*, decltype(&yyjson_mut_doc_free), yyjson_mut_doc_free> doc(yyjson_mut_doc_new(nullptr));
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }

    yyjson_mut_val* root = yyjson_mut_obj(doc.get());
    if (! root)
    {
        return E_OUTOFMEMORY;
    }
    yyjson_mut_doc_set_root(doc.get(), root);

    if (! yyjson_mut_obj_add_uint(doc.get(), root, "formatVersion", theme.formatVersion))
    {
        return E_OUTOFMEMORY;
    }

    const std::string idUtf8 = Utf8FromUtf16(theme.id);
    if (idUtf8.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (const HRESULT hr = AddStringMember(doc.get(), root, "id", idUtf8); FAILED(hr))
    {
        return hr;
    }

    const std::string nameUtf8 = Utf8FromUtf16(theme.name);
    if (nameUtf8.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (const HRESULT hr = AddStringMember(doc.get(), root, "name", nameUtf8); FAILED(hr))
    {
        return hr;
    }

    const std::string baseUtf8 = Utf8FromUtf16(theme.baseThemeId);
    if (baseUtf8.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (const HRESULT hr = AddStringMember(doc.get(), root, "baseThemeId", baseUtf8); FAILED(hr))
    {
        return hr;
    }

    const auto addSources = [&](const char* memberName, const std::unordered_map<std::wstring, ThemeColorSource>& sources, bool palette) noexcept -> HRESULT
    {
        yyjson_mut_val* object = yyjson_mut_obj(doc.get());
        if (! object || ! yyjson_mut_obj_add_val(doc.get(), root, memberName, object))
        {
            return E_OUTOFMEMORY;
        }
        std::vector<std::wstring_view> keys;
        keys.reserve(sources.size());
        for (const auto& [key, _] : sources)
        {
            if ((palette && ! IsValidThemePaletteName(key)) || (! palette && ! IsValidThemeColorKey(key)))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            keys.emplace_back(key);
        }
        std::sort(keys.begin(), keys.end());
        for (const std::wstring_view key : keys)
        {
            const auto it = sources.find(std::wstring(key));
            if (it == sources.end())
            {
                continue;
            }
            const std::string keyUtf8   = Utf8FromUtf16(key);
            const std::string valueUtf8 = Utf8FromUtf16(FormatThemeColorSource(it->second));
            if (keyUtf8.empty() || valueUtf8.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            yyjson_mut_val* keyNode   = yyjson_mut_strncpy(doc.get(), keyUtf8.data(), keyUtf8.size());
            yyjson_mut_val* valueNode = yyjson_mut_strncpy(doc.get(), valueUtf8.data(), valueUtf8.size());
            if (! keyNode || ! valueNode || ! yyjson_mut_obj_add(object, keyNode, valueNode))
            {
                return E_OUTOFMEMORY;
            }
        }
        return S_OK;
    };

    if (! theme.palette.empty())
    {
        if (const HRESULT hr = addSources("palette", theme.palette, true); FAILED(hr))
        {
            return hr;
        }
    }
    if (const HRESULT hr = addSources("colors", theme.colors, false); FAILED(hr))
    {
        return hr;
    }

    size_t jsonLength = 0;
    yyjson_write_err writeError{};
    const yyjson_write_flag flags = YYJSON_WRITE_PRETTY_TWO_SPACES | YYJSON_WRITE_NEWLINE_AT_END;
    wil::unique_any<char*, decltype(&std::free), std::free> jsonText(yyjson_mut_write_opts(doc.get(), flags, nullptr, &jsonLength, &writeError));
    if (! jsonText)
    {
        return writeError.code == YYJSON_WRITE_ERROR_MEMORY_ALLOCATION ? E_OUTOFMEMORY : E_FAIL;
    }

    outJson.assign(jsonText.get(), jsonLength);
    if (const size_t paletteOffset = outJson.find("  \"palette\": {"); paletteOffset != std::string::npos)
    {
        outJson.insert(paletteOffset, "  // Named palette\n");
    }

    const size_t colorsOffset = outJson.find("  \"colors\": {");
    if (colorsOffset != std::string::npos)
    {
        const size_t colorsBody = outJson.find('\n', colorsOffset);
        if (colorsBody != std::string::npos)
        {
            std::string grouped;
            grouped.reserve(outJson.size() + (theme.colors.size() * 12u));
            grouped.append(outJson.substr(0u, colorsBody + 1u));
            std::string previousGroup;
            size_t lineStart = colorsBody + 1u;
            while (lineStart < outJson.size())
            {
                size_t lineEnd = outJson.find('\n', lineStart);
                if (lineEnd == std::string::npos)
                    lineEnd = outJson.size();
                const std::string_view line(outJson.data() + lineStart, lineEnd - lineStart);
                if (line.starts_with("    \"") && ! line.starts_with("    //"))
                {
                    const size_t keyEnd = line.find('"', 5u);
                    if (keyEnd != std::string_view::npos)
                    {
                        const std::string_view key = line.substr(5u, keyEnd - 5u);
                        const size_t dot           = key.find('.');
                        const std::string group(key.substr(0u, dot));
                        if (group != previousGroup)
                        {
                            grouped.append("    // ");
                            grouped.append(group);
                            grouped.push_back('\n');
                            previousGroup = group;
                        }
                    }
                }
                grouped.append(line);
                if (lineEnd < outJson.size())
                    grouped.push_back('\n');
                lineStart = lineEnd + 1u;
            }
            outJson = std::move(grouped);
        }
    }
    return S_OK;
}
} // namespace Common::Settings
