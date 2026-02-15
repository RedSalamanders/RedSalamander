#pragma once

#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace SettingsSchemaParser
{

// Represents a single setting field extracted from the schema
struct SettingField
{
    std::wstring jsonPath;      // JSON path like "mainMenu.menuBarVisible"
    std::wstring paneName;      // Target pane: "General", "Advanced", "Keyboard", etc.
    std::wstring title;         // Display title from schema
    std::wstring description;   // Help text description
    std::wstring controlType;   // "toggle", "edit", "number", "combo", "custom"
    std::wstring sectionHeader; // Optional section grouping
    int displayOrder = 0;       // Sort order within pane

    // Type information from JSON Schema
    std::wstring schemaType; // "boolean", "string", "integer", "number", "array", "object"

    // For number types
    bool hasMin      = false;
    bool hasMax      = false;
    int64_t minValue = 0;
    int64_t maxValue = 0;

    // For enum/combo types
    std::vector<std::wstring> enumValues;

    // Default value (stored as string for flexibility)
    std::wstring defaultValue;
};

// Parse SettingsStore.schema.json and extract all fields with x-ui-pane attributes
// Returns fields sorted by pane → section → order
[[nodiscard]] std::vector<SettingField> ParseSettingsSchema(std::string_view schemaJsonUtf8) noexcept;

// Load schema from file and parse it
[[nodiscard]] std::vector<SettingField> LoadAndParseSettingsSchema(std::wstring_view schemaFilePath) noexcept;

// Filter fields by pane name
[[nodiscard]] std::vector<SettingField> GetFieldsForPane(const std::vector<SettingField>& allFields, std::wstring_view paneName) noexcept;

// Get only non-custom fields for a pane (excludes x-ui-control: custom)
[[nodiscard]] std::vector<SettingField> GetNonCustomFieldsForPane(const std::vector<SettingField>& allFields, std::wstring_view paneName) noexcept;

} // namespace SettingsSchemaParser
