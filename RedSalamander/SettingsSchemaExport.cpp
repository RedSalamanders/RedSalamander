#include "Framework.h"

#include "SettingsSchemaExport.h"

#include <algorithm>
#include <format>
#include <limits>
#include <new>
#include <unordered_set>

#include "FileSystemPluginManager.h"
#include "Helpers.h"
#include "ViewerPluginManager.h"

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL: deleted copy/move operators
#include <wil/resource.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

namespace
{
[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] uint32_t StableHash32(std::string_view text) noexcept
{
    // FNV-1a
    uint32_t hash = 2166136261u;
    for (const char ch : text)
    {
        hash ^= static_cast<unsigned char>(ch);
        hash *= 16777619u;
    }
    return hash;
}

[[nodiscard]] std::string MakePluginSchemaDefName(std::wstring_view pluginId) noexcept
{
    const std::string utf8 = Utf8FromUtf16(pluginId);
    const uint32_t hash    = StableHash32(utf8);

    std::string safe;
    safe.reserve(utf8.size());
    for (const char ch : utf8)
    {
        const bool ok = (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9');
        safe.push_back(ok ? ch : '_');
    }

    return std::format("pluginConfig_{}_{:08X}", safe, hash);
}

[[nodiscard]] yyjson_mut_val* BuildPluginConfigJsonSchema(yyjson_mut_doc* doc, std::wstring_view pluginId, std::string_view pluginSchemaJsonUtf8) noexcept
{
    yyjson_mut_val* schema = yyjson_mut_obj(doc);
    if (! schema)
    {
        return nullptr;
    }
    yyjson_mut_obj_add_strcpy(doc, schema, "type", "object");

    std::string pluginIdUtf8 = Utf8FromUtf16(pluginId);
    if (pluginIdUtf8.empty() && ! pluginId.empty())
    {
        yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", true);
        return schema;
    }

    if (pluginSchemaJsonUtf8.empty())
    {
        yyjson_mut_obj_add_strncpy(doc, schema, "title", pluginIdUtf8.c_str(), pluginIdUtf8.size());
        yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", true);
        return schema;
    }

    std::string mutableSchema;
    mutableSchema.assign(pluginSchemaJsonUtf8.data(), pluginSchemaJsonUtf8.size());

    yyjson_read_err err{};
    yyjson_doc* pluginDoc = yyjson_read_opts(mutableSchema.data(), mutableSchema.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err);
    if (! pluginDoc)
    {
        yyjson_mut_obj_add_strncpy(doc, schema, "title", pluginIdUtf8.c_str(), pluginIdUtf8.size());
        yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", true);
        return schema;
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(pluginDoc); });

    yyjson_val* root = yyjson_doc_get_root(pluginDoc);
    if (! root || ! yyjson_is_obj(root))
    {
        yyjson_mut_obj_add_strncpy(doc, schema, "title", pluginIdUtf8.c_str(), pluginIdUtf8.size());
        yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", true);
        return schema;
    }

    std::string_view titleText;
    if (yyjson_val* titleVal = yyjson_obj_get(root, "title"); titleVal && yyjson_is_str(titleVal))
    {
        const char* s = yyjson_get_str(titleVal);
        if (s)
        {
            titleText = std::string_view(s);
        }
    }

    if (! titleText.empty())
    {
        yyjson_mut_obj_add_strncpy(doc, schema, "title", titleText.data(), titleText.size());
    }
    else
    {
        yyjson_mut_obj_add_strncpy(doc, schema, "title", pluginIdUtf8.c_str(), pluginIdUtf8.size());
    }

    yyjson_val* fields = yyjson_obj_get(root, "fields");
    if (! fields || ! yyjson_is_arr(fields))
    {
        yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", true);
        return schema;
    }

    yyjson_mut_val* properties = yyjson_mut_obj(doc);
    if (! properties)
    {
        yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", true);
        return schema;
    }

    yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", false);
    yyjson_mut_obj_add_val(doc, schema, "properties", properties);

    const size_t count = yyjson_arr_size(fields);
    for (size_t i = 0; i < count; ++i)
    {
        yyjson_val* item = yyjson_arr_get(fields, i);
        if (! item || ! yyjson_is_obj(item))
        {
            continue;
        }

        yyjson_val* keyVal  = yyjson_obj_get(item, "key");
        yyjson_val* typeVal = yyjson_obj_get(item, "type");
        if (! keyVal || ! yyjson_is_str(keyVal) || ! typeVal || ! yyjson_is_str(typeVal))
        {
            continue;
        }

        const char* keyText  = yyjson_get_str(keyVal);
        const char* typeText = yyjson_get_str(typeVal);
        if (! keyText || ! typeText)
        {
            continue;
        }

        const std::string_view key(keyText);
        if (key.empty())
        {
            continue;
        }

        yyjson_mut_val* prop = yyjson_mut_obj(doc);
        if (! prop)
        {
            yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", true);
            return schema;
        }

        if (yyjson_val* labelVal = yyjson_obj_get(item, "label"); labelVal && yyjson_is_str(labelVal))
        {
            const char* labelText = yyjson_get_str(labelVal);
            if (labelText && labelText[0] != '\0')
            {
                yyjson_mut_obj_add_strcpy(doc, prop, "title", labelText);
            }
        }

        if (yyjson_val* descVal = yyjson_obj_get(item, "description"); descVal && yyjson_is_str(descVal))
        {
            const char* descText = yyjson_get_str(descVal);
            if (descText && descText[0] != '\0')
            {
                yyjson_mut_obj_add_strcpy(doc, prop, "description", descText);
            }
        }

        const std::string_view type(typeText);
        if (type == "text")
        {
            yyjson_mut_obj_add_strcpy(doc, prop, "type", "string");
            if (yyjson_val* defVal = yyjson_obj_get(item, "default"); defVal && yyjson_is_str(defVal))
            {
                const char* defText = yyjson_get_str(defVal);
                if (defText)
                {
                    yyjson_mut_obj_add_strcpy(doc, prop, "default", defText);
                }
            }
        }
        else if (type == "value")
        {
            yyjson_mut_obj_add_strcpy(doc, prop, "type", "integer");
            if (yyjson_val* defVal = yyjson_obj_get(item, "default"); defVal && yyjson_is_int(defVal))
            {
                yyjson_mut_obj_add_int(doc, prop, "default", yyjson_get_int(defVal));
            }
            if (yyjson_val* minVal = yyjson_obj_get(item, "min"); minVal && yyjson_is_int(minVal))
            {
                yyjson_mut_obj_add_int(doc, prop, "minimum", yyjson_get_int(minVal));
            }
            if (yyjson_val* maxVal = yyjson_obj_get(item, "max"); maxVal && yyjson_is_int(maxVal))
            {
                yyjson_mut_obj_add_int(doc, prop, "maximum", yyjson_get_int(maxVal));
            }
        }
        else if (type == "bool" || type == "boolean")
        {
            yyjson_mut_obj_add_strcpy(doc, prop, "type", "boolean");
            if (yyjson_val* defVal = yyjson_obj_get(item, "default"); defVal && yyjson_is_bool(defVal))
            {
                yyjson_mut_obj_add_bool(doc, prop, "default", yyjson_get_bool(defVal));
            }
        }
        else if (type == "option")
        {
            yyjson_mut_obj_add_strcpy(doc, prop, "type", "string");

            yyjson_mut_val* enumArr = nullptr;
            if (yyjson_val* options = yyjson_obj_get(item, "options"); options && yyjson_is_arr(options))
            {
                const size_t optCount = yyjson_arr_size(options);
                for (size_t o = 0; o < optCount; ++o)
                {
                    yyjson_val* opt = yyjson_arr_get(options, o);
                    if (! opt || ! yyjson_is_obj(opt))
                    {
                        continue;
                    }

                    yyjson_val* valueVal = yyjson_obj_get(opt, "value");
                    if (! valueVal || ! yyjson_is_str(valueVal))
                    {
                        continue;
                    }

                    const char* valueText = yyjson_get_str(valueVal);
                    if (! valueText || valueText[0] == '\0')
                    {
                        continue;
                    }

                    if (! enumArr)
                    {
                        enumArr = yyjson_mut_arr(doc);
                        if (! enumArr)
                        {
                            yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", true);
                            return schema;
                        }
                    }

                    yyjson_mut_arr_add_strcpy(doc, enumArr, valueText);
                }
            }

            if (enumArr)
            {
                yyjson_mut_obj_add_val(doc, prop, "enum", enumArr);
            }

            if (yyjson_val* defVal = yyjson_obj_get(item, "default"); defVal && yyjson_is_str(defVal))
            {
                const char* defText = yyjson_get_str(defVal);
                if (defText)
                {
                    yyjson_mut_obj_add_strcpy(doc, prop, "default", defText);
                }
            }
        }
        else if (type == "selection")
        {
            yyjson_mut_obj_add_strcpy(doc, prop, "type", "array");
            yyjson_mut_obj_add_bool(doc, prop, "uniqueItems", true);

            yyjson_mut_val* itemsSchema = yyjson_mut_obj(doc);
            if (! itemsSchema)
            {
                yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", true);
                return schema;
            }
            yyjson_mut_obj_add_strcpy(doc, itemsSchema, "type", "string");

            yyjson_mut_val* enumArr = nullptr;
            if (yyjson_val* options = yyjson_obj_get(item, "options"); options && yyjson_is_arr(options))
            {
                const size_t optCount = yyjson_arr_size(options);
                for (size_t o = 0; o < optCount; ++o)
                {
                    yyjson_val* opt = yyjson_arr_get(options, o);
                    if (! opt || ! yyjson_is_obj(opt))
                    {
                        continue;
                    }

                    yyjson_val* valueVal = yyjson_obj_get(opt, "value");
                    if (! valueVal || ! yyjson_is_str(valueVal))
                    {
                        continue;
                    }

                    const char* valueText = yyjson_get_str(valueVal);
                    if (! valueText || valueText[0] == '\0')
                    {
                        continue;
                    }

                    if (! enumArr)
                    {
                        enumArr = yyjson_mut_arr(doc);
                        if (! enumArr)
                        {
                            yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", true);
                            return schema;
                        }
                    }

                    yyjson_mut_arr_add_strcpy(doc, enumArr, valueText);
                }
            }

            if (enumArr)
            {
                yyjson_mut_obj_add_val(doc, itemsSchema, "enum", enumArr);
            }

            yyjson_mut_obj_add_val(doc, prop, "items", itemsSchema);

            if (yyjson_val* defVal = yyjson_obj_get(item, "default"); defVal && yyjson_is_arr(defVal))
            {
                yyjson_mut_val* defArr = yyjson_mut_arr(doc);
                if (! defArr)
                {
                    yyjson_mut_obj_add_bool(doc, schema, "additionalProperties", true);
                    return schema;
                }
                const size_t defCount = yyjson_arr_size(defVal);
                for (size_t d = 0; d < defCount; ++d)
                {
                    yyjson_val* v = yyjson_arr_get(defVal, d);
                    if (! v || ! yyjson_is_str(v))
                    {
                        continue;
                    }

                    const char* t = yyjson_get_str(v);
                    if (t && t[0] != '\0')
                    {
                        yyjson_mut_arr_add_strcpy(doc, defArr, t);
                    }
                }

                yyjson_mut_obj_add_val(doc, prop, "default", defArr);
            }
        }
        else
        {
            yyjson_mut_obj_add_bool(doc, prop, "additionalProperties", true);
        }

        yyjson_mut_val* keyOut = yyjson_mut_strncpy(doc, key.data(), key.size());
        if (! keyOut)
        {
            continue;
        }
        yyjson_mut_obj_add(properties, keyOut, prop);
    }

    return schema;
}

HRESULT BuildAggregatedSettingsSchemaJson(std::wstring_view appId, std::span<const PluginConfigurationSchemaSource> pluginSchemas, std::string& out) noexcept
{
    out.clear();

    const std::string_view baseSchemaJson = Common::Settings::GetSettingsStoreSchemaJsonUtf8();
    if (baseSchemaJson.empty())
    {
        return E_FAIL;
    }

    std::string mutableSchema;
    mutableSchema.assign(baseSchemaJson.data(), baseSchemaJson.size());

    yyjson_read_err baseErr{};
    yyjson_doc* baseDoc = yyjson_read_opts(mutableSchema.data(), mutableSchema.size(), YYJSON_READ_ALLOW_BOM, nullptr, &baseErr);
    if (! baseDoc)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    auto freeBase = wil::scope_exit([&] { yyjson_doc_free(baseDoc); });

    yyjson_mut_doc* doc = yyjson_doc_mut_copy(baseDoc, nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_doc_get_root(doc);
    if (! root || ! yyjson_mut_is_obj(root))
    {
        return E_FAIL;
    }

    const std::wstring commentWide = std::format(L"Generated by {} (aggregated plugin config schemas).", appId);
    const std::string commentUtf8  = Utf8FromUtf16(commentWide);
    if (! commentUtf8.empty())
    {
        yyjson_mut_obj_add_strncpy(doc, root, "$comment", commentUtf8.c_str(), commentUtf8.size());
    }

    yyjson_mut_val* defs = yyjson_mut_obj_get(root, "$defs");
    if (! defs || ! yyjson_mut_is_obj(defs))
    {
        return E_FAIL;
    }

    yyjson_mut_val* pluginsSettings = yyjson_mut_obj_get(defs, "pluginsSettings");
    if (! pluginsSettings || ! yyjson_mut_is_obj(pluginsSettings))
    {
        return E_FAIL;
    }

    yyjson_mut_val* pluginsProps = yyjson_mut_obj_get(pluginsSettings, "properties");
    if (! pluginsProps || ! yyjson_mut_is_obj(pluginsProps))
    {
        return E_FAIL;
    }

    yyjson_mut_val* configById = yyjson_mut_obj_get(pluginsProps, "configurationByPluginId");
    if (! configById || ! yyjson_mut_is_obj(configById))
    {
        return E_FAIL;
    }

    yyjson_mut_val* configProps = yyjson_mut_obj_get(configById, "properties");
    if (! configProps)
    {
        configProps = yyjson_mut_obj(doc);
        if (! configProps)
        {
            return E_OUTOFMEMORY;
        }
        if (! yyjson_mut_obj_add_val(doc, configById, "properties", configProps))
        {
            return E_OUTOFMEMORY;
        }
    }

    std::unordered_set<std::wstring> addedIds;

    for (const auto& plugin : pluginSchemas)
    {
        if (plugin.pluginId.empty())
        {
            continue;
        }

        if (addedIds.contains(plugin.pluginId))
        {
            continue;
        }

        const std::string pluginIdUtf8 = Utf8FromUtf16(plugin.pluginId);
        if (pluginIdUtf8.empty())
        {
            continue;
        }

        const std::string defName = MakePluginSchemaDefName(plugin.pluginId);
        yyjson_mut_val* schemaVal = BuildPluginConfigJsonSchema(doc, plugin.pluginId, plugin.schemaJsonUtf8);
        if (! schemaVal)
        {
            return E_OUTOFMEMORY;
        }

        yyjson_mut_val* defKey = yyjson_mut_strncpy(doc, defName.c_str(), defName.size());
        if (! defKey)
        {
            return E_OUTOFMEMORY;
        }

        if (! yyjson_mut_obj_add(defs, defKey, schemaVal))
        {
            return E_OUTOFMEMORY;
        }

        yyjson_mut_val* refObj = yyjson_mut_obj(doc);
        if (! refObj)
        {
            return E_OUTOFMEMORY;
        }
        const std::string ref = std::format("#/$defs/{}", defName);
        if (! yyjson_mut_obj_add_strncpy(doc, refObj, "$ref", ref.c_str(), ref.size()))
        {
            return E_OUTOFMEMORY;
        }

        yyjson_mut_val* propKey = yyjson_mut_strncpy(doc, pluginIdUtf8.c_str(), pluginIdUtf8.size());
        if (! propKey)
        {
            return E_OUTOFMEMORY;
        }

        if (! yyjson_mut_obj_add(configProps, propKey, refObj))
        {
            return E_OUTOFMEMORY;
        }
        static_cast<void>(addedIds.emplace(plugin.pluginId));
    }

    yyjson_write_err writeErr{};
    size_t len                    = 0;
    const yyjson_write_flag flags = YYJSON_WRITE_PRETTY_TWO_SPACES | YYJSON_WRITE_NEWLINE_AT_END;
    char* json                    = yyjson_mut_write_opts(doc, flags, nullptr, &len, &writeErr);
    if (! json)
    {
        return writeErr.code == YYJSON_WRITE_ERROR_MEMORY_ALLOCATION ? E_OUTOFMEMORY : E_FAIL;
    }

    auto freeJson = wil::scope_exit([&] { free(json); });

    if (len == 0)
    {
        return E_FAIL;
    }

    out.assign(json, len);
    return S_OK;
}
} // namespace

std::vector<PluginConfigurationSchemaSource> CollectPluginConfigurationSchemas(Common::Settings::Settings& settings) noexcept
{
    std::vector<PluginConfigurationSchemaSource> out;

    std::unordered_set<std::wstring> seen;

    auto tryAdd = [&](std::wstring_view pluginId, std::string schemaJsonUtf8)
    {
        if (pluginId.empty() || schemaJsonUtf8.empty())
        {
            return;
        }

        std::wstring id(pluginId);
        if (! seen.emplace(id).second)
        {
            return;
        }

        PluginConfigurationSchemaSource entry;
        entry.pluginId       = std::move(id);
        entry.schemaJsonUtf8 = std::move(schemaJsonUtf8);
        out.push_back(std::move(entry));
    };

    {
        auto& fs = FileSystemPluginManager::GetInstance();
        for (const auto& plugin : fs.GetPlugins())
        {
            if (plugin.id.empty())
            {
                continue;
            }

            std::string schema;
            if (SUCCEEDED(fs.GetConfigurationSchema(plugin.id, settings, schema)))
            {
                tryAdd(plugin.id, std::move(schema));
            }
        }
    }

    {
        auto& viewers = ViewerPluginManager::GetInstance();
        for (const auto& plugin : viewers.GetPlugins())
        {
            if (plugin.id.empty())
            {
                continue;
            }

            std::string schema;
            if (SUCCEEDED(viewers.GetConfigurationSchema(plugin.id, settings, schema)))
            {
                tryAdd(plugin.id, std::move(schema));
            }
        }
    }

    std::sort(
        out.begin(), out.end(), [](const PluginConfigurationSchemaSource& a, const PluginConfigurationSchemaSource& b) { return a.pluginId < b.pluginId; });

    return out;
}

HRESULT SaveAggregatedSettingsSchema(std::wstring_view appId, std::span<const PluginConfigurationSchemaSource> pluginSchemas) noexcept
{
    std::string schemaJsonUtf8;
    const HRESULT hr = BuildAggregatedSettingsSchemaJson(appId, pluginSchemas, schemaJsonUtf8);
    if (FAILED(hr))
    {
        return hr;
    }

    return Common::Settings::SaveSettingsSchema(appId, schemaJsonUtf8);
}

HRESULT SaveAggregatedSettingsSchema(std::wstring_view appId, Common::Settings::Settings& settings) noexcept
{
    const auto pluginSchemas = CollectPluginConfigurationSchemas(settings);
    return SaveAggregatedSettingsSchema(appId, pluginSchemas);
}
