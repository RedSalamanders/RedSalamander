#pragma once

#include <span>
#include <string>
#include <string_view>
#include <vector>

#include "SettingsStore.h"

struct PluginConfigurationSchemaSource
{
    std::wstring pluginId;
    std::string schemaJsonUtf8; // Plugin configuration schema (JSON/JSON5, UTF-8).
};

// Collects configuration schemas for all currently discovered plugins.
// Returns an empty vector if no plugins are available.
[[nodiscard]] std::vector<PluginConfigurationSchemaSource> CollectPluginConfigurationSchemas(Common::Settings::Settings& settings) noexcept;

// Writes an aggregated Settings Store JSON Schema alongside the settings file:
// - Base schema: Common::Settings::GetSettingsStoreSchemaJsonUtf8()
// - Plugin configuration schemas: converted from each plugin's `GetConfigurationSchema()` payload.
//
// The output path is `Common::Settings::GetSettingsSchemaPath(appId)`.
HRESULT SaveAggregatedSettingsSchema(std::wstring_view appId, std::span<const PluginConfigurationSchemaSource> pluginSchemas) noexcept;

// Convenience wrapper: collects plugin schemas from currently loaded/discovered plugin managers, then writes the aggregated schema.
HRESULT SaveAggregatedSettingsSchema(std::wstring_view appId, Common::Settings::Settings& settings) noexcept;
