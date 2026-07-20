#pragma once

#include <cstddef>
#include <cstdint>
#include <span>
#include <string>
#include <string_view>
#include <vector>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#ifndef COMMON_API
#ifdef COMMON_EXPORTS
#define COMMON_API __declspec(dllexport)
#else
#define COMMON_API __declspec(dllimport)
#endif
#endif

namespace Common::PluginConfiguration
{
enum class FieldType : uint8_t
{
    Text,
    Value,
    Bool,
    Option,
    Selection,
};

struct Choice
{
    std::wstring value;
    std::wstring label;
};

struct Field
{
    FieldType type = FieldType::Text;
    std::wstring key;
    std::wstring label;
    std::wstring description;
    bool browseFolder = false;

    bool hasMin = false;
    bool hasMax = false;
    int64_t min = 0;
    int64_t max = 0;

    std::wstring defaultText;
    int64_t defaultInt = 0;
    bool defaultBool   = false;
    std::wstring defaultOption;
    std::vector<std::wstring> defaultSelection;
    std::vector<Choice> choices;

    std::wstring uiSection;
    int uiOrder = 0;
    std::wstring uiControl;
    bool uiHidden = false;
};

enum class ValidationSeverity : uint8_t
{
    Warning,
    Error,
};

enum class ValidationCode : uint8_t
{
    InvalidJson,
    RootNotObject,
    FieldsMissingOrInvalid,
    FieldNotObject,
    MissingOrInvalidKey,
    MissingOrInvalidType,
    UnsupportedType,
    DuplicateFieldKey,
    InvalidConstraintRange,
    InvalidDefault,
    InvalidOptions,
    DuplicateOptionValue,
    ConfigurationValueWrongType,
};

inline constexpr size_t kNoFieldIndex = (static_cast<size_t>(-1));

struct ValidationIssue
{
    ValidationSeverity severity = ValidationSeverity::Warning;
    ValidationCode code         = ValidationCode::InvalidJson;
    size_t fieldIndex           = kNoFieldIndex;
    std::wstring fieldKey;
};

struct SchemaParseResult
{
    std::vector<Field> fields;
    std::vector<ValidationIssue> issues;

    [[nodiscard]] bool HasErrors() const noexcept
    {
        for (const ValidationIssue& issue : issues)
        {
            if (issue.severity == ValidationSeverity::Error)
            {
                return true;
            }
        }
        return false;
    }
};

struct FieldValue
{
    FieldType type = FieldType::Text;
    std::wstring text;
    int64_t integer = 0;
    bool boolean    = false;
    std::vector<std::wstring> selection;
};

struct ConfigurationParseResult
{
    std::vector<FieldValue> values;
    std::vector<ValidationIssue> issues;
    bool sourceWasObject = false;

    [[nodiscard]] bool HasErrors() const noexcept
    {
        for (const ValidationIssue& issue : issues)
        {
            if (issue.severity == ValidationSeverity::Error)
            {
                return true;
            }
        }
        return false;
    }
};

COMMON_API SchemaParseResult ParseSchema(std::string_view schemaJsonUtf8) noexcept;
COMMON_API ConfigurationParseResult ParseConfiguration(std::span<const Field> fields, std::string_view configurationJsonUtf8) noexcept;
COMMON_API bool TryGetBoolToggleChoiceIndices(const Field& field, size_t& outOnIndex, size_t& outOffIndex) noexcept;

// Serializes current values over the original object. Members that are not represented by the schema are
// copied unchanged and in their original order. Known members retain their original position when updated.
COMMON_API HRESULT SerializeConfiguration(std::string_view originalConfigurationJsonUtf8,
                                          std::span<const Field> fields,
                                          std::span<const FieldValue> values,
                                          std::string& outConfigurationJsonUtf8) noexcept;
} // namespace Common::PluginConfiguration
