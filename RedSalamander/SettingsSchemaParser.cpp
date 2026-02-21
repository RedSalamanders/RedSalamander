#include "Framework.h"

#include "SettingsSchemaParser.h"

#include <algorithm>
#include <filesystem>
#include <fstream>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

namespace SettingsSchemaParser
{

namespace
{

// UTF-8 <-> UTF-16 conversion helpers
[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::optional<std::string_view> TryGetUtf8String(yyjson_val* obj, const char* key) noexcept
{
    if (! obj || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* val = yyjson_obj_get(obj, key);
    if (! val || ! yyjson_is_str(val))
    {
        return std::nullopt;
    }

    const char* str = yyjson_get_str(val);
    if (! str)
    {
        return std::nullopt;
    }

    return std::string_view(str);
}

[[nodiscard]] std::optional<int64_t> TryGetInt64(yyjson_val* obj, const char* key) noexcept
{
    if (! obj || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* val = yyjson_obj_get(obj, key);
    if (! val || ! yyjson_is_int(val))
    {
        return std::nullopt;
    }

    return yyjson_get_sint(val);
}

// Recursively walk JSON schema properties and extract fields with x-ui-pane
void WalkSchemaProperties(yyjson_val* propertiesObj, const std::wstring& currentPath, std::vector<SettingField>& outFields) noexcept
{
    if (! propertiesObj || ! yyjson_is_obj(propertiesObj))
    {
        return;
    }

    // Iterate over all properties
    yyjson_obj_iter iter;
    yyjson_obj_iter_init(propertiesObj, &iter);
    yyjson_val* key = nullptr;

    while ((key = yyjson_obj_iter_next(&iter)) != nullptr)
    {
        const char* keyName = yyjson_get_str(key);
        if (! keyName)
        {
            continue;
        }

        yyjson_val* propValue = yyjson_obj_iter_get_val(key);
        if (! propValue || ! yyjson_is_obj(propValue))
        {
            continue;
        }

        // Build JSON path
        std::wstring jsonPath = currentPath;
        if (! jsonPath.empty())
        {
            jsonPath += L".";
        }
        jsonPath += Utf16FromUtf8(keyName);

        // Check if this property has x-ui-pane attribute
        const auto uiPane = TryGetUtf8String(propValue, "x-ui-pane");
        if (uiPane.has_value())
        {
            // This is a UI field - extract its metadata
            SettingField field;
            field.jsonPath = jsonPath;
            field.paneName = Utf16FromUtf8(uiPane.value());

            // Extract x-ui-* attributes
            const auto uiControl = TryGetUtf8String(propValue, "x-ui-control");
            field.controlType    = uiControl.has_value() ? Utf16FromUtf8(uiControl.value()) : L"edit";

            const auto uiSection = TryGetUtf8String(propValue, "x-ui-section");
            field.sectionHeader  = uiSection.has_value() ? Utf16FromUtf8(uiSection.value()) : L"";

            const auto uiOrder = TryGetInt64(propValue, "x-ui-order");
            field.displayOrder = uiOrder.has_value() ? static_cast<int>(uiOrder.value()) : 0;

            // Extract schema metadata
            const auto title = TryGetUtf8String(propValue, "title");
            field.title      = title.has_value() ? Utf16FromUtf8(title.value()) : jsonPath;

            const auto description = TryGetUtf8String(propValue, "description");
            field.description      = description.has_value() ? Utf16FromUtf8(description.value()) : L"";

            const auto schemaType = TryGetUtf8String(propValue, "type");
            field.schemaType      = schemaType.has_value() ? Utf16FromUtf8(schemaType.value()) : L"string";

            // Extract min/max for numeric types
            const auto minVal = TryGetInt64(propValue, "minimum");
            if (minVal.has_value())
            {
                field.hasMin   = true;
                field.minValue = minVal.value();
            }

            const auto maxVal = TryGetInt64(propValue, "maximum");
            if (maxVal.has_value())
            {
                field.hasMax   = true;
                field.maxValue = maxVal.value();
            }

            // Extract enum values
            yyjson_val* enumArray = yyjson_obj_get(propValue, "enum");
            if (enumArray && yyjson_is_arr(enumArray))
            {
                const size_t enumCount = yyjson_arr_size(enumArray);
                field.enumValues.reserve(enumCount);

                for (size_t i = 0; i < enumCount; ++i)
                {
                    yyjson_val* enumVal = yyjson_arr_get(enumArray, i);
                    if (enumVal && yyjson_is_str(enumVal))
                    {
                        const char* enumStr = yyjson_get_str(enumVal);
                        if (enumStr)
                        {
                            field.enumValues.push_back(Utf16FromUtf8(enumStr));
                        }
                    }
                }
            }

            // Extract default value
            yyjson_val* defaultVal = yyjson_obj_get(propValue, "default");
            if (defaultVal)
            {
                if (yyjson_is_bool(defaultVal))
                {
                    field.defaultValue = yyjson_get_bool(defaultVal) ? L"true" : L"false";
                }
                else if (yyjson_is_int(defaultVal))
                {
                    field.defaultValue = std::to_wstring(yyjson_get_sint(defaultVal));
                }
                else if (yyjson_is_str(defaultVal))
                {
                    const char* str = yyjson_get_str(defaultVal);
                    if (str)
                    {
                        field.defaultValue = Utf16FromUtf8(str);
                    }
                }
            }

            outFields.push_back(std::move(field));
        }

        // Recursively process nested properties
        yyjson_val* nestedProps = yyjson_obj_get(propValue, "properties");
        if (nestedProps && yyjson_is_obj(nestedProps))
        {
            WalkSchemaProperties(nestedProps, jsonPath, outFields);
        }
    }
}

// Walk $defs section and extract fields
void WalkSchemaDefinitions(yyjson_val* defsObj, std::vector<SettingField>& outFields) noexcept
{
    if (! defsObj || ! yyjson_is_obj(defsObj))
    {
        return;
    }

    // Iterate over all definitions
    yyjson_obj_iter iter;
    yyjson_obj_iter_init(defsObj, &iter);
    yyjson_val* key = nullptr;

    while ((key = yyjson_obj_iter_next(&iter)) != nullptr)
    {
        yyjson_val* defValue = yyjson_obj_iter_get_val(key);
        if (! defValue || ! yyjson_is_obj(defValue))
        {
            continue;
        }

        // Check if this definition itself has x-ui-pane (for top-level settings like themeSettings)
        const auto uiPane = TryGetUtf8String(defValue, "x-ui-pane");
        if (uiPane.has_value())
        {
            const char* defName = yyjson_get_str(key);
            if (defName)
            {
                SettingField field;
                field.jsonPath = Utf16FromUtf8(defName);
                field.paneName = Utf16FromUtf8(uiPane.value());

                const auto uiControl = TryGetUtf8String(defValue, "x-ui-control");
                field.controlType    = uiControl.has_value() ? Utf16FromUtf8(uiControl.value()) : L"custom";

                const auto uiOrder = TryGetInt64(defValue, "x-ui-order");
                field.displayOrder = uiOrder.has_value() ? static_cast<int>(uiOrder.value()) : 0;

                const auto title = TryGetUtf8String(defValue, "title");
                field.title      = title.has_value() ? Utf16FromUtf8(title.value()) : field.jsonPath;

                const auto description = TryGetUtf8String(defValue, "description");
                field.description      = description.has_value() ? Utf16FromUtf8(description.value()) : L"";

                outFields.push_back(std::move(field));
            }
        }

        // Process nested properties within this definition
        yyjson_val* nestedProps = yyjson_obj_get(defValue, "properties");
        if (nestedProps && yyjson_is_obj(nestedProps))
        {
            const char* defName         = yyjson_get_str(key);
            const std::wstring basePath = defName ? Utf16FromUtf8(defName) : L"";
            WalkSchemaProperties(nestedProps, basePath, outFields);
        }
    }
}

} // anonymous namespace

std::vector<SettingField> ParseSettingsSchema(std::string_view schemaJsonUtf8) noexcept
{
    std::vector<SettingField> fields;

    if (schemaJsonUtf8.empty())
    {
        return fields;
    }

    yyjson_doc* doc = yyjson_read(schemaJsonUtf8.data(), schemaJsonUtf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
    if (! doc)
    {
        return fields;
    }

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        yyjson_doc_free(doc);
        return fields;
    }

    // Process top-level properties
    yyjson_val* properties = yyjson_obj_get(root, "properties");
    if (properties && yyjson_is_obj(properties))
    {
        WalkSchemaProperties(properties, L"", fields);
    }

    // Process $defs section
    yyjson_val* defs = yyjson_obj_get(root, "$defs");
    if (defs && yyjson_is_obj(defs))
    {
        WalkSchemaDefinitions(defs, fields);
    }

    yyjson_doc_free(doc);

    // Sort by pane → section → order
    std::sort(fields.begin(),
              fields.end(),
              [](const SettingField& a, const SettingField& b)
              {
                  if (a.paneName != b.paneName)
                  {
                      return a.paneName < b.paneName;
                  }
                  if (a.sectionHeader != b.sectionHeader)
                  {
                      return a.sectionHeader < b.sectionHeader;
                  }
                  return a.displayOrder < b.displayOrder;
              });

    return fields;
}

std::vector<SettingField> LoadAndParseSettingsSchema(std::wstring_view schemaFilePath) noexcept
{
    std::vector<SettingField> fields;

    std::ifstream file(schemaFilePath.data(), std::ios::binary);
    if (! file.is_open())
    {
        return fields;
    }

    // Read entire file
    file.seekg(0, std::ios::end);
    const auto fileSize = file.tellg();
    file.seekg(0, std::ios::beg);

    if (fileSize <= 0 || fileSize > 10 * 1024 * 1024) // Max 10MB
    {
        return fields;
    }

    std::string content(static_cast<size_t>(fileSize), '\0');
    file.read(content.data(), fileSize);

    if (file.gcount() != fileSize)
    {
        return fields;
    }

    return ParseSettingsSchema(content);
}

std::vector<SettingField> GetFieldsForPane(const std::vector<SettingField>& allFields, std::wstring_view paneName) noexcept
{
    std::vector<SettingField> result;
    result.reserve(allFields.size() / 4); // Rough estimate

    for (const auto& field : allFields)
    {
        if (field.paneName == paneName)
        {
            result.push_back(field);
        }
    }

    return result;
}

std::vector<SettingField> GetNonCustomFieldsForPane(const std::vector<SettingField>& allFields, std::wstring_view paneName) noexcept
{
    std::vector<SettingField> result;
    result.reserve(allFields.size() / 4);

    for (const auto& field : allFields)
    {
        if (field.paneName == paneName && field.controlType != L"custom")
        {
            result.push_back(field);
        }
    }

    return result;
}

} // namespace SettingsSchemaParser
