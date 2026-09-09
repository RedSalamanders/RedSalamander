#include "PluginConfiguration.h"

#include "Helpers.h"
#include "SettingsStore.h"
#include "YyjsonHelpers.h"

#include <algorithm>
#include <limits>
#include <optional>
#include <unordered_set>
#include <utility>

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

namespace Common::PluginConfiguration
{
namespace
{
void AddIssue(
    std::vector<ValidationIssue>& issues, ValidationSeverity severity, ValidationCode code, size_t fieldIndex = kNoFieldIndex, std::wstring fieldKey = {})
{
    issues.push_back(ValidationIssue{.severity = severity, .code = code, .fieldIndex = fieldIndex, .fieldKey = std::move(fieldKey)});
}

[[nodiscard]] std::optional<std::string_view> GetString(yyjson_val* object, const char* key) noexcept
{
    if (! object || ! key)
    {
        return std::nullopt;
    }
    yyjson_val* value = yyjson_obj_get(object, key);
    if (! value || ! yyjson_is_str(value))
    {
        return std::nullopt;
    }
    const char* text = yyjson_get_str(value);
    return text ? std::make_optional(std::string_view(text, yyjson_get_len(value))) : std::nullopt;
}

[[nodiscard]] bool GetInt64(yyjson_val* object, const char* key, int64_t& out) noexcept
{
    yyjson_val* value = object && key ? yyjson_obj_get(object, key) : nullptr;
    if (! value)
    {
        return false;
    }
    if (yyjson_is_sint(value))
    {
        out = yyjson_get_sint(value);
        return true;
    }
    if (yyjson_is_uint(value))
    {
        out = static_cast<int64_t>((std::min)(yyjson_get_uint(value), static_cast<uint64_t>((std::numeric_limits<int64_t>::max)())));
        return true;
    }
    if (yyjson_is_real(value))
    {
        const double real = yyjson_get_real(value);
        if (real <= static_cast<double>((std::numeric_limits<int64_t>::min)()))
        {
            out = (std::numeric_limits<int64_t>::min)();
        }
        else if (real >= static_cast<double>((std::numeric_limits<int64_t>::max)()))
        {
            out = (std::numeric_limits<int64_t>::max)();
        }
        else
        {
            out = static_cast<int64_t>(real);
        }
        return true;
    }
    return false;
}

[[nodiscard]] bool EqualsNoCase(std::wstring_view lhs, std::wstring_view rhs) noexcept
{
    return OrdinalString::EqualsNoCase(lhs, rhs);
}

[[nodiscard]] std::optional<bool> ParseBoolToken(std::wstring_view token) noexcept
{
    if (EqualsNoCase(token, L"on") || EqualsNoCase(token, L"true") || token == L"1")
    {
        return true;
    }
    if (EqualsNoCase(token, L"off") || EqualsNoCase(token, L"false") || token == L"0")
    {
        return false;
    }
    return std::nullopt;
}

[[nodiscard]] bool GetBool(yyjson_val* object, const char* key, bool& out) noexcept
{
    yyjson_val* value = object && key ? yyjson_obj_get(object, key) : nullptr;
    if (! value)
    {
        return false;
    }
    if (yyjson_is_bool(value))
    {
        out = yyjson_get_bool(value);
        return true;
    }
    if (yyjson_is_sint(value))
    {
        out = yyjson_get_sint(value) != 0;
        return true;
    }
    if (yyjson_is_uint(value))
    {
        out = yyjson_get_uint(value) != 0;
        return true;
    }
    if (yyjson_is_str(value))
    {
        const char* text = yyjson_get_str(value);
        if (text)
        {
            const std::optional<bool> parsed = ParseBoolToken(Common::Strings::Utf16FromUtf8StrictOrEmpty(text));
            if (parsed.has_value())
            {
                out = parsed.value();
                return true;
            }
        }
    }
    return false;
}

[[nodiscard]] std::optional<FieldType> ParseFieldType(std::string_view type) noexcept
{
    if (type == "text")
    {
        return FieldType::Text;
    }
    if (type == "value")
    {
        return FieldType::Value;
    }
    if (type == "bool" || type == "boolean")
    {
        return FieldType::Bool;
    }
    if (type == "option")
    {
        return FieldType::Option;
    }
    if (type == "selection")
    {
        return FieldType::Selection;
    }
    return std::nullopt;
}

void ParseChoices(yyjson_val* item, size_t fieldIndex, Field& field, std::vector<ValidationIssue>& issues)
{
    yyjson_val* options = yyjson_obj_get(item, "options");
    if (! options || ! yyjson_is_arr(options))
    {
        AddIssue(issues, ValidationSeverity::Warning, ValidationCode::InvalidOptions, fieldIndex, field.key);
        return;
    }

    std::unordered_set<std::wstring> values;
    const size_t optionCount = yyjson_arr_size(options);
    field.choices.reserve(optionCount);
    for (size_t optionIndex = 0; optionIndex < optionCount; ++optionIndex)
    {
        yyjson_val* option                              = yyjson_arr_get(options, optionIndex);
        const std::optional<std::string_view> valueUtf8 = option && yyjson_is_obj(option) ? GetString(option, "value") : std::nullopt;
        if (! valueUtf8.has_value())
        {
            AddIssue(issues, ValidationSeverity::Warning, ValidationCode::InvalidOptions, fieldIndex, field.key);
            continue;
        }

        Choice choice;
        choice.value = Common::Strings::Utf16FromUtf8StrictOrEmpty(valueUtf8.value());
        if (choice.value.empty())
        {
            AddIssue(issues, ValidationSeverity::Warning, ValidationCode::InvalidOptions, fieldIndex, field.key);
            continue;
        }
        if (! values.emplace(choice.value).second)
        {
            AddIssue(issues, ValidationSeverity::Warning, ValidationCode::DuplicateOptionValue, fieldIndex, field.key);
            continue;
        }

        const std::optional<std::string_view> labelUtf8 = GetString(option, "label");
        choice.label                                    = labelUtf8.has_value() ? Common::Strings::Utf16FromUtf8StrictOrEmpty(labelUtf8.value()) : choice.value;
        if (choice.label.empty())
        {
            choice.label = choice.value;
        }
        field.choices.push_back(std::move(choice));
    }
}

[[nodiscard]] FieldValue MakeDefaultValue(const Field& field)
{
    FieldValue value;
    value.type = field.type;
    switch (field.type)
    {
        case FieldType::Text: value.text = field.defaultText; break;
        case FieldType::Value: value.integer = field.defaultInt; break;
        case FieldType::Bool: value.boolean = field.defaultBool; break;
        case FieldType::Option: value.text = field.defaultOption; break;
        case FieldType::Selection: value.selection = field.defaultSelection; break;
    }
    return value;
}

[[nodiscard]] Settings::JsonValue CloneJsonValue(const Settings::JsonValue& source)
{
    Settings::JsonValue clone;
    if (const auto* array = std::get_if<Settings::JsonValue::ArrayPtr>(&source.value))
    {
        auto target = std::make_shared<Settings::JsonArray>();
        if (*array)
        {
            target->items.reserve((*array)->items.size());
            for (const Settings::JsonValue& item : (*array)->items)
            {
                target->items.push_back(CloneJsonValue(item));
            }
        }
        clone.value = std::move(target);
    }
    else if (const auto* object = std::get_if<Settings::JsonValue::ObjectPtr>(&source.value))
    {
        auto target = std::make_shared<Settings::JsonObject>();
        if (*object)
        {
            target->members.reserve((*object)->members.size());
            for (const auto& [key, value] : (*object)->members)
            {
                target->members.emplace_back(key, CloneJsonValue(value));
            }
        }
        clone.value = std::move(target);
    }
    else
    {
        clone.value = source.value;
    }
    return clone;
}

[[nodiscard]] Settings::JsonValue ToJsonValue(const Field& field, const FieldValue& value)
{
    Settings::JsonValue result;
    switch (field.type)
    {
        case FieldType::Text:
        case FieldType::Option: result.value = Common::Strings::Utf8FromUtf16StrictOrEmpty(value.text); break;
        case FieldType::Value:
        {
            int64_t integer = value.integer;
            if (field.hasMin)
            {
                integer = (std::max)(integer, field.min);
            }
            if (field.hasMax)
            {
                integer = (std::min)(integer, field.max);
            }
            result.value = integer;
            break;
        }
        case FieldType::Bool: result.value = value.boolean; break;
        case FieldType::Selection:
        {
            auto array = std::make_shared<Settings::JsonArray>();
            array->items.reserve(value.selection.size());
            for (const std::wstring& selected : value.selection)
            {
                Settings::JsonValue item;
                item.value = Common::Strings::Utf8FromUtf16StrictOrEmpty(selected);
                array->items.push_back(std::move(item));
            }
            result.value = std::move(array);
            break;
        }
    }
    return result;
}
} // namespace

bool TryGetBoolToggleChoiceIndices(const Field& field, size_t& outOnIndex, size_t& outOffIndex) noexcept
{
    if (field.type != FieldType::Option || field.choices.size() != 2)
    {
        return false;
    }

    std::optional<size_t> onIndex;
    std::optional<size_t> offIndex;
    for (size_t choiceIndex = 0; choiceIndex < field.choices.size(); ++choiceIndex)
    {
        const Choice& choice       = field.choices[choiceIndex];
        std::optional<bool> parsed = ParseBoolToken(choice.label);
        if (! parsed.has_value())
        {
            parsed = ParseBoolToken(choice.value);
        }
        if (! parsed.has_value())
        {
            continue;
        }

        if (parsed.value())
        {
            onIndex = choiceIndex;
        }
        else
        {
            offIndex = choiceIndex;
        }
    }

    if (! onIndex.has_value() || ! offIndex.has_value() || onIndex.value() == offIndex.value())
    {
        return false;
    }

    outOnIndex  = onIndex.value();
    outOffIndex = offIndex.value();
    return true;
}

SchemaParseResult ParseSchema(std::string_view schemaJsonUtf8) noexcept
{
    SchemaParseResult result;
    if (schemaJsonUtf8.empty())
    {
        AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::InvalidJson);
        return result;
    }

    yyjson_doc* rawDocument = yyjson_read(schemaJsonUtf8.data(), schemaJsonUtf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
    if (! rawDocument)
    {
        AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::InvalidJson);
        return result;
    }
    Common::Json::UniqueDocument document{rawDocument};
    yyjson_val* root = yyjson_doc_get_root(document.get());
    if (! root || ! yyjson_is_obj(root))
    {
        AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::RootNotObject);
        return result;
    }

    yyjson_val* fieldsArray = yyjson_obj_get(root, "fields");
    if (! fieldsArray || ! yyjson_is_arr(fieldsArray))
    {
        AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::FieldsMissingOrInvalid);
        return result;
    }

    const size_t fieldCount = yyjson_arr_size(fieldsArray);
    result.fields.reserve(fieldCount);
    std::unordered_set<std::wstring> fieldKeys;
    for (size_t fieldIndex = 0; fieldIndex < fieldCount; ++fieldIndex)
    {
        yyjson_val* item = yyjson_arr_get(fieldsArray, fieldIndex);
        if (! item || ! yyjson_is_obj(item))
        {
            AddIssue(result.issues, ValidationSeverity::Warning, ValidationCode::FieldNotObject, fieldIndex);
            continue;
        }

        const std::optional<std::string_view> keyUtf8 = GetString(item, "key");
        if (! keyUtf8.has_value())
        {
            AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::MissingOrInvalidKey, fieldIndex);
            continue;
        }
        std::wstring key = Common::Strings::Utf16FromUtf8StrictOrEmpty(keyUtf8.value());
        if (key.empty())
        {
            AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::MissingOrInvalidKey, fieldIndex);
            continue;
        }
        if (! fieldKeys.emplace(key).second)
        {
            AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::DuplicateFieldKey, fieldIndex, std::move(key));
            continue;
        }

        const std::optional<std::string_view> typeUtf8 = GetString(item, "type");
        if (! typeUtf8.has_value())
        {
            AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::MissingOrInvalidType, fieldIndex, std::move(key));
            continue;
        }
        const std::optional<FieldType> fieldType = ParseFieldType(typeUtf8.value());
        if (! fieldType.has_value())
        {
            AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::UnsupportedType, fieldIndex, std::move(key));
            continue;
        }

        Field field;
        field.type                                      = fieldType.value();
        field.key                                       = std::move(key);
        const std::optional<std::string_view> labelUtf8 = GetString(item, "label");
        field.label                                     = labelUtf8.has_value() ? Common::Strings::Utf16FromUtf8StrictOrEmpty(labelUtf8.value()) : field.key;
        if (field.label.empty())
        {
            field.label = field.key;
        }
        if (const std::optional<std::string_view> descriptionUtf8 = GetString(item, "description"); descriptionUtf8.has_value())
        {
            field.description = Common::Strings::Utf16FromUtf8StrictOrEmpty(descriptionUtf8.value());
        }
        if (const std::optional<std::string_view> browseUtf8 = GetString(item, "browse"); browseUtf8.has_value())
        {
            field.browseFolder = browseUtf8.value() == "folder" || browseUtf8.value() == "directory";
        }
        if (const std::optional<std::string_view> sectionUtf8 = GetString(item, "x-ui-section"); sectionUtf8.has_value())
        {
            field.uiSection = Common::Strings::Utf16FromUtf8StrictOrEmpty(sectionUtf8.value());
        }
        if (const std::optional<std::string_view> controlUtf8 = GetString(item, "x-ui-control"); controlUtf8.has_value())
        {
            field.uiControl = Common::Strings::Utf16FromUtf8StrictOrEmpty(controlUtf8.value());
        }
        bool hidden = false;
        if (GetBool(item, "x-ui-hidden", hidden))
        {
            field.uiHidden = hidden;
        }
        int64_t order = 0;
        if (GetInt64(item, "x-ui-order", order))
        {
            field.uiOrder = static_cast<int>(
                std::clamp<int64_t>(order, static_cast<int64_t>((std::numeric_limits<int>::min)()), static_cast<int64_t>((std::numeric_limits<int>::max)())));
        }

        if (GetInt64(item, "min", field.min))
        {
            field.hasMin = true;
        }
        if (GetInt64(item, "max", field.max))
        {
            field.hasMax = true;
        }
        if (field.hasMin && field.hasMax && field.min > field.max)
        {
            AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::InvalidConstraintRange, fieldIndex, field.key);
        }

        switch (field.type)
        {
            case FieldType::Text:
                if (const std::optional<std::string_view> value = GetString(item, "default"); value.has_value())
                {
                    field.defaultText = Common::Strings::Utf16FromUtf8StrictOrEmpty(value.value());
                }
                else if (yyjson_obj_get(item, "default"))
                {
                    AddIssue(result.issues, ValidationSeverity::Warning, ValidationCode::InvalidDefault, fieldIndex, field.key);
                }
                break;
            case FieldType::Value:
                if (! GetInt64(item, "default", field.defaultInt) && yyjson_obj_get(item, "default"))
                {
                    AddIssue(result.issues, ValidationSeverity::Warning, ValidationCode::InvalidDefault, fieldIndex, field.key);
                }
                if ((field.hasMin && field.defaultInt < field.min) || (field.hasMax && field.defaultInt > field.max))
                {
                    AddIssue(result.issues, ValidationSeverity::Warning, ValidationCode::InvalidDefault, fieldIndex, field.key);
                }
                break;
            case FieldType::Bool:
                if (! GetBool(item, "default", field.defaultBool) && yyjson_obj_get(item, "default"))
                {
                    AddIssue(result.issues, ValidationSeverity::Warning, ValidationCode::InvalidDefault, fieldIndex, field.key);
                }
                break;
            case FieldType::Option:
                if (const std::optional<std::string_view> value = GetString(item, "default"); value.has_value())
                {
                    field.defaultOption = Common::Strings::Utf16FromUtf8StrictOrEmpty(value.value());
                }
                else if (yyjson_obj_get(item, "default"))
                {
                    AddIssue(result.issues, ValidationSeverity::Warning, ValidationCode::InvalidDefault, fieldIndex, field.key);
                }
                ParseChoices(item, fieldIndex, field, result.issues);
                if (! field.defaultOption.empty() &&
                    std::ranges::none_of(field.choices, [&](const Choice& choice) noexcept { return choice.value == field.defaultOption; }))
                {
                    AddIssue(result.issues, ValidationSeverity::Warning, ValidationCode::InvalidDefault, fieldIndex, field.key);
                }
                break;
            case FieldType::Selection:
            {
                ParseChoices(item, fieldIndex, field, result.issues);
                yyjson_val* defaults = yyjson_obj_get(item, "default");
                if (defaults && yyjson_is_arr(defaults))
                {
                    const size_t defaultCount = yyjson_arr_size(defaults);
                    field.defaultSelection.reserve(defaultCount);
                    for (size_t defaultIndex = 0; defaultIndex < defaultCount; ++defaultIndex)
                    {
                        yyjson_val* value = yyjson_arr_get(defaults, defaultIndex);
                        if (! value || ! yyjson_is_str(value))
                        {
                            AddIssue(result.issues, ValidationSeverity::Warning, ValidationCode::InvalidDefault, fieldIndex, field.key);
                            continue;
                        }
                        const char* text     = yyjson_get_str(value);
                        std::wstring decoded = text ? Common::Strings::Utf16FromUtf8StrictOrEmpty(text) : std::wstring{};
                        if (decoded.empty())
                        {
                            AddIssue(result.issues, ValidationSeverity::Warning, ValidationCode::InvalidDefault, fieldIndex, field.key);
                            continue;
                        }
                        field.defaultSelection.push_back(std::move(decoded));
                    }
                }
                else if (defaults)
                {
                    AddIssue(result.issues, ValidationSeverity::Warning, ValidationCode::InvalidDefault, fieldIndex, field.key);
                }
                break;
            }
        }
        result.fields.push_back(std::move(field));
    }

    std::stable_sort(result.fields.begin(),
                     result.fields.end(),
                     [](const Field& lhs, const Field& rhs) noexcept
    {
        if (lhs.uiOrder != 0 && rhs.uiOrder != 0)
        {
            return lhs.uiOrder < rhs.uiOrder;
        }
        if (lhs.uiOrder != 0)
        {
            return true;
        }
        return false;
    });
    return result;
}

ConfigurationParseResult ParseConfiguration(std::span<const Field> fields, std::string_view configurationJsonUtf8) noexcept
{
    ConfigurationParseResult result;
    result.values.reserve(fields.size());
    for (const Field& field : fields)
    {
        result.values.push_back(MakeDefaultValue(field));
    }
    if (configurationJsonUtf8.empty())
    {
        return result;
    }

    Settings::JsonValue root;
    if (FAILED(Settings::ParseJsonValue(configurationJsonUtf8, root)))
    {
        AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::InvalidJson);
        return result;
    }
    const auto* object = std::get_if<Settings::JsonValue::ObjectPtr>(&root.value);
    if (! object || ! *object)
    {
        AddIssue(result.issues, ValidationSeverity::Error, ValidationCode::RootNotObject);
        return result;
    }
    result.sourceWasObject = true;

    for (size_t fieldIndex = 0; fieldIndex < fields.size(); ++fieldIndex)
    {
        const Field& field                 = fields[fieldIndex];
        FieldValue& value                  = result.values[fieldIndex];
        const std::string keyUtf8          = Common::Strings::Utf8FromUtf16StrictOrEmpty(field.key);
        const Settings::JsonValue* current = keyUtf8.empty() ? nullptr : Settings::FindMember(root, keyUtf8);
        if (! current)
        {
            continue;
        }

        bool valid = false;
        switch (field.type)
        {
            case FieldType::Text:
            case FieldType::Option:
                if (const auto* text = std::get_if<std::string>(&current->value))
                {
                    std::wstring decoded = Common::Strings::Utf16FromUtf8StrictOrEmpty(*text);
                    if (decoded.empty() && ! text->empty())
                    {
                        break;
                    }
                    value.text = std::move(decoded);
                    valid      = true;
                }
                break;
            case FieldType::Value:
                if (const auto* integer = std::get_if<int64_t>(&current->value))
                {
                    value.integer = *integer;
                    valid         = true;
                }
                else if (const auto* unsignedInteger = std::get_if<uint64_t>(&current->value))
                {
                    value.integer = static_cast<int64_t>((std::min)(*unsignedInteger, static_cast<uint64_t>((std::numeric_limits<int64_t>::max)())));
                    valid         = true;
                }
                else if (const auto* real = std::get_if<double>(&current->value))
                {
                    if (*real <= static_cast<double>((std::numeric_limits<int64_t>::min)()))
                    {
                        value.integer = (std::numeric_limits<int64_t>::min)();
                    }
                    else if (*real >= static_cast<double>((std::numeric_limits<int64_t>::max)()))
                    {
                        value.integer = (std::numeric_limits<int64_t>::max)();
                    }
                    else
                    {
                        value.integer = static_cast<int64_t>(*real);
                    }
                    valid = true;
                }
                break;
            case FieldType::Bool:
                if (const auto* boolean = std::get_if<bool>(&current->value))
                {
                    value.boolean = *boolean;
                    valid         = true;
                }
                else if (const auto* integer = std::get_if<int64_t>(&current->value))
                {
                    value.boolean = *integer != 0;
                    valid         = true;
                }
                else if (const auto* unsignedInteger = std::get_if<uint64_t>(&current->value))
                {
                    value.boolean = *unsignedInteger != 0u;
                    valid         = true;
                }
                else if (const auto* text = std::get_if<std::string>(&current->value))
                {
                    const std::optional<bool> parsed = ParseBoolToken(Common::Strings::Utf16FromUtf8StrictOrEmpty(*text));
                    if (parsed.has_value())
                    {
                        value.boolean = parsed.value();
                        valid         = true;
                    }
                }
                break;
            case FieldType::Selection:
                if (const auto* array = std::get_if<Settings::JsonValue::ArrayPtr>(&current->value); array && *array)
                {
                    value.selection.clear();
                    for (const Settings::JsonValue& item : (*array)->items)
                    {
                        const auto* text = std::get_if<std::string>(&item.value);
                        if (! text)
                        {
                            continue;
                        }
                        std::wstring decoded = Common::Strings::Utf16FromUtf8StrictOrEmpty(*text);
                        if (! decoded.empty())
                        {
                            value.selection.push_back(std::move(decoded));
                        }
                    }
                    valid = true;
                }
                break;
        }
        if (! valid)
        {
            AddIssue(result.issues, ValidationSeverity::Warning, ValidationCode::ConfigurationValueWrongType, fieldIndex, field.key);
        }
    }
    return result;
}

HRESULT SerializeConfiguration(std::string_view originalConfigurationJsonUtf8,
                               std::span<const Field> fields,
                               std::span<const FieldValue> values,
                               std::string& outConfigurationJsonUtf8) noexcept
{
    outConfigurationJsonUtf8.clear();
    if (fields.size() != values.size())
    {
        return E_INVALIDARG;
    }

    Settings::JsonValue root;
    bool recoveredSource = false;
    if (! originalConfigurationJsonUtf8.empty())
    {
        const HRESULT parseHr = Settings::ParseJsonValue(originalConfigurationJsonUtf8, root);
        if (parseHr == E_OUTOFMEMORY)
        {
            return parseHr;
        }
        recoveredSource = FAILED(parseHr);
    }
    const auto* originalObject = std::get_if<Settings::JsonValue::ObjectPtr>(&root.value);
    if (! originalObject || ! *originalObject)
    {
        recoveredSource = recoveredSource || ! originalConfigurationJsonUtf8.empty();
        root.value      = std::make_shared<Settings::JsonObject>();
    }
    else
    {
        root = CloneJsonValue(root);
    }

    auto* object = std::get_if<Settings::JsonValue::ObjectPtr>(&root.value);
    if (! object || ! *object)
    {
        return E_OUTOFMEMORY;
    }

    for (size_t fieldIndex = 0; fieldIndex < fields.size(); ++fieldIndex)
    {
        const Field& field      = fields[fieldIndex];
        const FieldValue& value = values[fieldIndex];
        if (field.type != value.type)
        {
            return E_INVALIDARG;
        }
        const std::string keyUtf8 = Common::Strings::Utf8FromUtf16StrictOrEmpty(field.key);
        if (keyUtf8.empty())
        {
            return E_INVALIDARG;
        }

        Settings::JsonValue serialized = ToJsonValue(field, value);
        auto first                     = std::ranges::find_if((*object)->members, [&](const auto& member) noexcept { return member.first == keyUtf8; });
        if (first == (*object)->members.end())
        {
            (*object)->members.emplace_back(keyUtf8, std::move(serialized));
            continue;
        }
        first->second = std::move(serialized);
        (*object)->members.erase(
            std::remove_if(std::next(first), (*object)->members.end(), [&](const auto& member) noexcept { return member.first == keyUtf8; }),
            (*object)->members.end());
    }

    const HRESULT serializeHr = Settings::SerializeJsonValue(root, outConfigurationJsonUtf8);
    if (FAILED(serializeHr))
    {
        return serializeHr;
    }
    return recoveredSource ? S_FALSE : S_OK;
}
} // namespace Common::PluginConfiguration
