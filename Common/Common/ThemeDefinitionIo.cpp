#include "ThemeDefinitionIo.h"

#include <algorithm>
#include <array>
#include <cstdlib>
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
    if (text.empty() || text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    const int required = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    const int required = ::WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        ::WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] yyjson_val* GetObjectMember(yyjson_val* obj, const char* key) noexcept
{
    yyjson_val* value = yyjson_obj_get(obj, key);
    return (value && yyjson_is_obj(value)) ? value : nullptr;
}

[[nodiscard]] HRESULT ReadRequiredString(yyjson_val* obj,
                                         const char* key,
                                         std::wstring& out,
                                         Common::Settings::ThemeDefinitionIoError missingError,
                                         Common::Settings::ThemeDefinitionIoError* outError,
                                         std::wstring* outMessage) noexcept
{
    out.clear();
    yyjson_val* value = yyjson_obj_get(obj, key);
    if (! value || ! yyjson_is_str(value))
    {
        return InvalidData(missingError, outError, outMessage);
    }

    const char* text = yyjson_get_str(value);
    if (! text)
    {
        return InvalidData(missingError, outError, outMessage);
    }

    out = Utf16FromUtf8(std::string_view(text, yyjson_get_len(value)));
    if (out.empty())
    {
        return InvalidData(missingError, outError, outMessage);
    }

    return S_OK;
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

HRESULT ParseThemeDefinitionJson5(std::string_view jsonText, ThemeDefinition& outTheme, ThemeDefinitionIoError* outError, std::wstring* outMessage) noexcept
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

    if (jsonText.empty())
    {
        return InvalidData(ThemeDefinitionIoError::EmptyInput, outError, outMessage);
    }

    std::string mutableJson(jsonText);
    yyjson_read_err readError{};
    wil::unique_any<yyjson_doc*, decltype(&yyjson_doc_free), yyjson_doc_free> doc(
        yyjson_read_opts(mutableJson.data(), mutableJson.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &readError));
    if (! doc)
    {
        if (readError.code == YYJSON_READ_ERROR_MEMORY_ALLOCATION)
        {
            return OutOfMemory(outError, outMessage);
        }

        const std::wstring message = (readError.msg && readError.msg[0] != '\0') ? Utf16FromUtf8(readError.msg) : std::wstring{};
        return InvalidData(ThemeDefinitionIoError::ParseFailed, outError, outMessage, message);
    }

    yyjson_val* root = yyjson_doc_get_root(doc.get());
    if (! root || ! yyjson_is_obj(root))
    {
        return InvalidData(ThemeDefinitionIoError::RootNotObject, outError, outMessage);
    }

    ThemeDefinition parsed;
    if (const HRESULT hr = ReadRequiredString(root, "id", parsed.id, ThemeDefinitionIoError::MissingOrInvalidId, outError, outMessage); FAILED(hr))
    {
        return hr;
    }
    if (! IsValidUserThemeId(parsed.id))
    {
        return InvalidData(ThemeDefinitionIoError::InvalidId, outError, outMessage);
    }

    if (const HRESULT hr = ReadRequiredString(root, "name", parsed.name, ThemeDefinitionIoError::MissingOrInvalidName, outError, outMessage); FAILED(hr))
    {
        return hr;
    }
    if (parsed.name.size() > 64u)
    {
        return InvalidData(ThemeDefinitionIoError::MissingOrInvalidName, outError, outMessage);
    }

    if (const HRESULT hr =
            ReadRequiredString(root, "baseThemeId", parsed.baseThemeId, ThemeDefinitionIoError::MissingOrInvalidBaseThemeId, outError, outMessage);
        FAILED(hr))
    {
        return hr;
    }
    if (! IsBuiltinThemeId(parsed.baseThemeId))
    {
        return InvalidData(ThemeDefinitionIoError::InvalidBaseThemeId, outError, outMessage);
    }

    yyjson_val* colors = GetObjectMember(root, "colors");
    if (! colors)
    {
        return InvalidData(ThemeDefinitionIoError::ColorsMissingOrNotObject, outError, outMessage);
    }

    yyjson_obj_iter iter = yyjson_obj_iter_with(colors);
    yyjson_val* colorKey = nullptr;
    while ((colorKey = yyjson_obj_iter_next(&iter)) != nullptr)
    {
        if (! yyjson_is_str(colorKey))
        {
            return InvalidData(ThemeDefinitionIoError::InvalidColorKey, outError, outMessage);
        }

        const char* keyText = yyjson_get_str(colorKey);
        if (! keyText)
        {
            return InvalidData(ThemeDefinitionIoError::InvalidColorKey, outError, outMessage);
        }

        const std::wstring key = Utf16FromUtf8(std::string_view(keyText, yyjson_get_len(colorKey)));
        if (! IsValidThemeColorKey(key))
        {
            return InvalidData(ThemeDefinitionIoError::InvalidColorKey, outError, outMessage);
        }

        yyjson_val* colorValue = yyjson_obj_iter_get_val(colorKey);
        if (! colorValue || ! yyjson_is_str(colorValue))
        {
            return InvalidData(ThemeDefinitionIoError::ColorValueNotString, outError, outMessage);
        }

        const char* valueText = yyjson_get_str(colorValue);
        if (! valueText)
        {
            return InvalidData(ThemeDefinitionIoError::InvalidColorValue, outError, outMessage);
        }

        const std::wstring value = Utf16FromUtf8(std::string_view(valueText, yyjson_get_len(colorValue)));
        uint32_t argb            = 0;
        if (value.empty() || ! TryParseColor(value, argb))
        {
            return InvalidData(ThemeDefinitionIoError::InvalidColorValue, outError, outMessage);
        }

        parsed.colors[key] = argb;
    }

    outTheme = std::move(parsed);
    return S_OK;
}

HRESULT BuildThemeDefinitionJson5(const ThemeDefinition& theme, std::string& outJson) noexcept
{
    outJson.clear();

    if (! IsValidUserThemeId(theme.id) || theme.name.empty() || theme.name.size() > 64u || ! IsBuiltinThemeId(theme.baseThemeId))
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

    yyjson_mut_val* colors = yyjson_mut_obj(doc.get());
    if (! colors || ! yyjson_mut_obj_add_val(doc.get(), root, "colors", colors))
    {
        return E_OUTOFMEMORY;
    }

    std::vector<std::wstring_view> keys;
    keys.reserve(theme.colors.size());
    for (const auto& [key, _] : theme.colors)
    {
        if (! IsValidThemeColorKey(key))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        keys.emplace_back(key);
    }

    std::sort(keys.begin(), keys.end());

    for (const std::wstring_view key : keys)
    {
        const auto it = theme.colors.find(std::wstring(key));
        if (it == theme.colors.end())
        {
            continue;
        }

        const std::string keyUtf8 = Utf8FromUtf16(key);
        if (keyUtf8.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const std::wstring colorText = FormatColor(it->second);
        const std::string valueUtf8  = Utf8FromUtf16(colorText);
        if (valueUtf8.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        yyjson_mut_val* keyNode   = yyjson_mut_strncpy(doc.get(), keyUtf8.data(), keyUtf8.size());
        yyjson_mut_val* valueNode = yyjson_mut_strncpy(doc.get(), valueUtf8.data(), valueUtf8.size());
        if (! keyNode || ! valueNode || ! yyjson_mut_obj_add(colors, keyNode, valueNode))
        {
            return E_OUTOFMEMORY;
        }
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
    return S_OK;
}
} // namespace Common::Settings
