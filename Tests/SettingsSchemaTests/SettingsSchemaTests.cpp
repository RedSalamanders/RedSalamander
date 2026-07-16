// SettingsSchemaTests - direct unit tests for SettingsSchemaParser.
//
// The parser is pure string-in / struct-out logic that turns
// Specs/SettingsStore.schema.json into the field model that drives the
// Preferences UI and settings persistence. These tests exercise it against the
// real shipped schema (characterization) and against synthetic inline JSON
// (defaults, min/max, enums, malformed input, missing file, non-ASCII).
//
// Pass/fail strings ("[       OK ]" / "[ FAILED  ]") and the wmain 0/1 contract
// match the other Tests\ suites so the runner's grep treats this like the rest.

#include <algorithm>
#include <filesystem>
#include <iostream>
#include <string>
#include <string_view>
#include <vector>

#include "PluginConfiguration.h"
#include "SettingsStore.h"
#include "SettingsSchemaParser.h"

namespace
{

// Copied verbatim from Tests\LocalizationTests\LocalizationTests.cpp:33-43 so the
// pass/fail bracket spacing matches across suites.
void Check(bool condition, const wchar_t* message, bool& success) noexcept
{
    if (! condition)
    {
        std::wcerr << L"[ FAILED  ] " << message << L"\n";
        success = false;
        return;
    }

    std::wcout << L"[       OK ] " << message << L"\n";
}

using SettingsSchemaParser::SettingField;

// Resolve Specs\SettingsStore.schema.json robustly: the exe runs from
// .build\x64\Debug, so walk up from the current directory looking for the file.
// argv[1] (if given) is preferred. Returns empty if not found.
[[nodiscard]] std::wstring FindSchemaPath(int argc, wchar_t** argv) noexcept
{
    namespace fs = std::filesystem;

    if (argc >= 2 && argv[1] != nullptr)
    {
        std::error_code ec;
        if (fs::exists(fs::path(argv[1]), ec))
        {
            return std::wstring(argv[1]);
        }
    }

    std::error_code ec;
    fs::path dir = fs::current_path(ec);
    if (ec)
    {
        return {};
    }

    for (int i = 0; i < 12; ++i)
    {
        const fs::path candidate = dir / L"Specs" / L"SettingsStore.schema.json";
        if (fs::exists(candidate, ec))
        {
            return candidate.wstring();
        }

        const fs::path parent = dir.parent_path();
        if (parent == dir)
        {
            break;
        }
        dir = parent;
    }

    return {};
}

[[nodiscard]] const SettingField* FindByPath(const std::vector<SettingField>& fields, std::wstring_view jsonPath) noexcept
{
    for (const auto& f : fields)
    {
        if (f.jsonPath == jsonPath)
        {
            return &f;
        }
    }
    return nullptr;
}

// --- Step 2: real-schema characterization -------------------------------------

void TestRealSchema(const std::wstring& schemaPath, bool& success)
{
    const std::vector<SettingField> fields = SettingsSchemaParser::LoadAndParseSettingsSchema(schemaPath);

    // 1. Non-empty.
    Check(! fields.empty(), L"real schema: parse returns a non-empty field set", success);
    if (fields.empty())
    {
        return; // nothing else is meaningful
    }

    // 2. Anchor fields (verified against Specs\SettingsStore.schema.json:1316-1340).
    const SettingField* menuBar = FindByPath(fields, L"mainMenuSettings.menuBarVisible");
    Check(menuBar != nullptr, L"real schema: anchor mainMenuSettings.menuBarVisible present", success);
    if (menuBar != nullptr)
    {
        Check(menuBar->paneName == L"General", L"anchor menuBarVisible: paneName == General", success);
        Check(menuBar->schemaType == L"boolean", L"anchor menuBarVisible: schemaType == boolean", success);
        Check(menuBar->controlType == L"toggle", L"anchor menuBarVisible: controlType == toggle", success);
        Check(menuBar->sectionHeader == L"Display", L"anchor menuBarVisible: section == Display", success);
        Check(menuBar->displayOrder == 10, L"anchor menuBarVisible: displayOrder == 10", success);
        Check(menuBar->defaultValue == L"true", L"anchor menuBarVisible: defaultValue == true", success);
    }

    const SettingField* funcBar = FindByPath(fields, L"mainMenuSettings.functionBarVisible");
    Check(funcBar != nullptr, L"real schema: anchor mainMenuSettings.functionBarVisible present", success);
    if (funcBar != nullptr)
    {
        Check(funcBar->paneName == L"General", L"anchor functionBarVisible: paneName == General", success);
        Check(funcBar->schemaType == L"boolean", L"anchor functionBarVisible: schemaType == boolean", success);
        Check(funcBar->controlType == L"toggle", L"anchor functionBarVisible: controlType == toggle", success);
        Check(funcBar->displayOrder == 20, L"anchor functionBarVisible: displayOrder == 20", success);
    }

    // 3. Sort invariant: grouped by pane, then section, then non-descending order
    //    within a pane+section (mirrors the parser's std::sort comparator).
    bool sortedOk = true;
    for (size_t i = 1; i < fields.size(); ++i)
    {
        const SettingField& a = fields[i - 1];
        const SettingField& b = fields[i];
        if (a.paneName != b.paneName)
        {
            if (! (a.paneName < b.paneName))
            {
                sortedOk = false;
                break;
            }
        }
        else if (a.sectionHeader != b.sectionHeader)
        {
            if (! (a.sectionHeader < b.sectionHeader))
            {
                sortedOk = false;
                break;
            }
        }
        else if (a.displayOrder > b.displayOrder)
        {
            sortedOk = false;
            break;
        }
    }
    Check(sortedOk, L"real schema: result sorted by pane -> section -> displayOrder", success);

    // 4. GetFieldsForPane: known pane returns only that pane and at least one;
    //    garbage pane returns empty.
    const std::vector<SettingField> general = SettingsSchemaParser::GetFieldsForPane(fields, L"General");
    Check(! general.empty(), L"GetFieldsForPane(General): returns at least one field", success);
    bool allGeneral = true;
    for (const auto& f : general)
    {
        if (f.paneName != L"General")
        {
            allGeneral = false;
            break;
        }
    }
    Check(allGeneral, L"GetFieldsForPane(General): every field belongs to General", success);

    const std::vector<SettingField> nonexistent = SettingsSchemaParser::GetFieldsForPane(fields, L"NoSuchPane");
    Check(nonexistent.empty(), L"GetFieldsForPane(NoSuchPane): returns empty", success);

    // 5. GetNonCustomFieldsForPane: never includes a custom field, AND for a pane
    //    that owns at least one custom field it returns strictly fewer than
    //    GetFieldsForPane (proving the filter drops something). Discover such a
    //    pane from the parsed data rather than hardcoding.
    std::wstring paneWithCustom;
    for (const auto& f : fields)
    {
        if (f.controlType == L"custom")
        {
            paneWithCustom = f.paneName;
            break;
        }
    }
    Check(! paneWithCustom.empty(), L"real schema: at least one custom-control field exists", success);
    if (! paneWithCustom.empty())
    {
        const std::vector<SettingField> all     = SettingsSchemaParser::GetFieldsForPane(fields, paneWithCustom);
        const std::vector<SettingField> noncust = SettingsSchemaParser::GetNonCustomFieldsForPane(fields, paneWithCustom);

        bool noCustomInResult = true;
        for (const auto& f : noncust)
        {
            if (f.controlType == L"custom")
            {
                noCustomInResult = false;
                break;
            }
        }
        Check(noCustomInResult, L"GetNonCustomFieldsForPane: result contains no custom fields", success);
        Check(noncust.size() < all.size(), L"GetNonCustomFieldsForPane: drops at least one custom field", success);
    }
}

// --- Step 3: synthetic inline-JSON tests --------------------------------------

void TestSyntheticDefaults(bool& success)
{
    // properties-level boolean field with x-ui-pane but NO x-ui-control.
    const char* json = R"({
        "type": "object",
        "properties": {
            "flag": {
                "type": "boolean",
                "default": true,
                "x-ui-pane": "General"
            }
        }
    })";

    const std::vector<SettingField> fields = SettingsSchemaParser::ParseSettingsSchema(json);
    Check(fields.size() == 1, L"synthetic defaults: exactly one field parsed", success);
    if (fields.size() == 1)
    {
        const SettingField& f = fields[0];
        Check(f.jsonPath == L"flag", L"synthetic defaults: jsonPath == flag", success);
        Check(f.controlType == L"edit", L"synthetic defaults: missing x-ui-control defaults to edit", success);
        Check(f.title == f.jsonPath, L"synthetic defaults: missing title defaults to jsonPath", success);
        Check(f.schemaType == L"boolean", L"synthetic defaults: schemaType == boolean", success);
        Check(f.defaultValue == L"true", L"synthetic defaults: bool default true round-trips to L\"true\"", success);
    }
}

void TestSyntheticMinMax(bool& success)
{
    const char* withBounds                  = R"({
        "properties": {
            "count": {
                "type": "integer",
                "minimum": 1,
                "maximum": 100,
                "x-ui-pane": "Advanced",
                "x-ui-control": "number"
            }
        }
    })";
    const std::vector<SettingField> bounded = SettingsSchemaParser::ParseSettingsSchema(withBounds);
    Check(bounded.size() == 1, L"synthetic min/max: one field with bounds", success);
    if (bounded.size() == 1)
    {
        Check(bounded[0].hasMin && bounded[0].minValue == 1, L"synthetic min/max: hasMin true, minValue 1", success);
        Check(bounded[0].hasMax && bounded[0].maxValue == 100, L"synthetic min/max: hasMax true, maxValue 100", success);
    }

    const char* noBounds                      = R"({
        "properties": {
            "count": {
                "type": "integer",
                "x-ui-pane": "Advanced",
                "x-ui-control": "number"
            }
        }
    })";
    const std::vector<SettingField> unbounded = SettingsSchemaParser::ParseSettingsSchema(noBounds);
    Check(unbounded.size() == 1, L"synthetic min/max: one field without bounds", success);
    if (unbounded.size() == 1)
    {
        Check(! unbounded[0].hasMin && ! unbounded[0].hasMax, L"synthetic min/max: no bounds -> hasMin/hasMax false", success);
    }
}

void TestSyntheticEnum(bool& success)
{
    const char* json                       = R"({
        "properties": {
            "mode": {
                "type": "string",
                "enum": ["alpha", "beta", "gamma"],
                "x-ui-pane": "General",
                "x-ui-control": "combo"
            }
        }
    })";
    const std::vector<SettingField> fields = SettingsSchemaParser::ParseSettingsSchema(json);
    Check(fields.size() == 1, L"synthetic enum: one combo field", success);
    if (fields.size() == 1)
    {
        const std::vector<std::wstring>& e = fields[0].enumValues;
        const bool ok                      = e.size() == 3 && e[0] == L"alpha" && e[1] == L"beta" && e[2] == L"gamma";
        Check(ok, L"synthetic enum: enumValues populated in schema order", success);
    }
}

void TestSyntheticNoPane(bool& success)
{
    // Field WITHOUT x-ui-pane must NOT be returned.
    const char* json                       = R"({
        "properties": {
            "hidden": {
                "type": "boolean",
                "default": false
            }
        }
    })";
    const std::vector<SettingField> fields = SettingsSchemaParser::ParseSettingsSchema(json);
    Check(fields.empty(), L"synthetic no-pane: field without x-ui-pane is not returned", success);
}

void TestMalformedInput(bool& success)
{
    // These are characterization checks: a noexcept parser is expected to return
    // an empty vector (or skip the bad field) rather than crash. If any of these
    // crashed instead, that is a STOP condition reported separately.

    Check(SettingsSchemaParser::ParseSettingsSchema("").empty(), L"malformed: empty string -> empty vector", success);

    // Truncated JSON.
    Check(SettingsSchemaParser::ParseSettingsSchema("{ \"properties\": { \"a\":").empty(), L"malformed: truncated JSON -> empty vector", success);

    // Root is not an object.
    Check(SettingsSchemaParser::ParseSettingsSchema("[1, 2, 3]").empty(), L"malformed: array root -> empty vector", success);

    // Wrong type for minimum ("abc"): the field is still pane-attributed, so it
    // is returned but minimum is ignored (TryGetInt64 rejects non-int).
    const char* badMin                  = R"({
        "properties": {
            "count": {
                "type": "integer",
                "minimum": "abc",
                "x-ui-pane": "Advanced",
                "x-ui-control": "number"
            }
        }
    })";
    const std::vector<SettingField> bad = SettingsSchemaParser::ParseSettingsSchema(badMin);
    // Documents current behavior: field is kept, bad minimum is simply not set.
    const bool badMinOk = bad.size() == 1 && ! bad[0].hasMin;
    Check(badMinOk, L"malformed: minimum:\"abc\" -> field kept, hasMin stays false", success);
}

void TestMissingFile(bool& success)
{
    const std::vector<SettingField> fields = SettingsSchemaParser::LoadAndParseSettingsSchema(L"Z:\\definitely\\no\\such\\schema_file_xyz.json");
    Check(fields.empty(), L"missing file: LoadAndParseSettingsSchema returns empty, no crash", success);
}

void TestNonAscii(bool& success)
{
    // UTF-8 -> wstring round trip for title/description (German + accented chars).
    // Build the UTF-8 bytes explicitly. Splitting adjacent string literals keeps
    // a trailing letter from being absorbed into a preceding \xNN escape
    // (e.g. "\xC3\x9F" "e" rather than "\xC3\x9Fe").
    const char* json                       = "{\n"
                                             "  \"properties\": {\n"
                                             "    \"locale\": {\n"
                                             "      \"type\": \"string\",\n"
                                             "      \"title\": \"Gr\xC3\xB6\xC3\x9F"
                                             "e\",\n" // "Größe"
                                             "      \"description\": \"\xC3\x89l\xC3\xA9"
                                             "ment\",\n" // "Élément"
                                             "      \"x-ui-pane\": \"General\",\n"
                                             "      \"x-ui-control\": \"edit\"\n"
                                             "    }\n"
                                             "  }\n"
                                             "}\n";
    const std::vector<SettingField> fields = SettingsSchemaParser::ParseSettingsSchema(json);
    Check(fields.size() == 1, L"non-ASCII: one field parsed", success);
    if (fields.size() == 1)
    {
        Check(fields[0].title == L"Größe", L"non-ASCII: title round-trips (Größe)", success);
        Check(fields[0].description == L"Élément", L"non-ASCII: description round-trips (Élément)", success);
    }
}

[[nodiscard]] bool HasPluginConfigurationIssue(const Common::PluginConfiguration::SchemaParseResult& result,
                                               Common::PluginConfiguration::ValidationCode code) noexcept
{
    return std::ranges::any_of(result.issues,
                               [code](const Common::PluginConfiguration::ValidationIssue& issue) noexcept { return issue.code == code; });
}

void TestPluginConfigurationModelAndCodec(bool& success)
{
    using namespace Common::PluginConfiguration;

    constexpr std::string_view schemaJson = R"json({
        "fields": [
            { "key": "text", "type": "text", "label": "Text", "description": "Text value", "browse": "folder", "default": "default" },
            { "key": "count", "type": "value", "min": 1, "max": 9, "default": 4, "x-ui-order": 20 },
            { "key": "enabled", "type": "boolean", "default": "on", "x-ui-order": 10 },
            { "key": "mode", "type": "option", "default": "a", "options": [
                { "value": "a", "label": "Alpha" }, { "value": "b" }
            ] },
            { "key": "tags", "type": "selection", "default": ["a"], "options": [
                { "value": "a" }, { "value": "b", "label": "Beta" }
            ] },
            { "key": "hidden", "type": "text", "default": "secret", "x-ui-hidden": true,
              "x-ui-section": "Advanced", "x-ui-control": "custom", "futureSchemaMember": { "x": 1 } }
        ],
        "futureRootMember": true
    })json";
    const SchemaParseResult schema = ParseSchema(schemaJson);
    Check(! schema.HasErrors() && schema.fields.size() == 6u, L"plugin config schema: every supported field type parses without errors", success);
    if (schema.fields.size() != 6u)
    {
        return;
    }

    Check(schema.fields[0].key == L"enabled" && schema.fields[1].key == L"count" && schema.fields[2].key == L"text",
          L"plugin config schema: explicit UI order sorts first and stable source order follows",
          success);
    const auto textField = std::ranges::find_if(schema.fields, [](const Field& field) noexcept { return field.key == L"text"; });
    Check(textField != schema.fields.end() && textField->browseFolder, L"plugin config schema: folder browse metadata is shared", success);
    const auto hiddenField = std::ranges::find_if(schema.fields, [](const Field& field) noexcept { return field.key == L"hidden"; });
    Check(hiddenField != schema.fields.end() && hiddenField->uiHidden && hiddenField->uiSection == L"Advanced" && hiddenField->uiControl == L"custom",
          L"plugin config schema: shared UI metadata is preserved",
          success);

    Field booleanOption;
    booleanOption.type    = FieldType::Option;
    booleanOption.choices = {{L"disabled", L"Off"}, {L"enabled", L"On"}};
    size_t onChoiceIndex  = kNoFieldIndex;
    size_t offChoiceIndex = kNoFieldIndex;
    Check(TryGetBoolToggleChoiceIndices(booleanOption, onChoiceIndex, offChoiceIndex) && onChoiceIndex == 1u && offChoiceIndex == 0u,
          L"plugin config schema: bool-like option labels share the same toggle mapping",
          success);
    booleanOption.choices = {{L"false", L"Disabled"}, {L"true", L"Enabled"}};
    Check(TryGetBoolToggleChoiceIndices(booleanOption, onChoiceIndex, offChoiceIndex) && onChoiceIndex == 1u && offChoiceIndex == 0u,
          L"plugin config schema: bool-like option values are a toggle fallback",
          success);
    booleanOption.choices.push_back({L"auto", L"Automatic"});
    Check(! TryGetBoolToggleChoiceIndices(booleanOption, onChoiceIndex, offChoiceIndex),
          L"plugin config schema: multi-choice options remain combo boxes",
          success);

    constexpr std::string_view configurationJson = R"json({
        "futureBefore": { "nested": [1, "keep"] },
        "text": "configured",
        "count": 7.9,
        "enabled": "off",
        "mode": "b",
        "tags": ["a", 7, "future-option"],
        "hidden": "retained-hidden",
        "futureAfter": true
    })json";
    ConfigurationParseResult configuration = ParseConfiguration(schema.fields, configurationJson);
    Check(! configuration.HasErrors() && configuration.sourceWasObject && configuration.values.size() == schema.fields.size(),
          L"plugin config model: configuration object overlays schema defaults",
          success);
    if (configuration.values.size() != schema.fields.size())
    {
        return;
    }

    const auto indexOf = [&](std::wstring_view key) noexcept
    {
        const auto field = std::ranges::find_if(schema.fields, [&](const Field& candidate) noexcept { return candidate.key == key; });
        return field == schema.fields.end() ? kNoFieldIndex : static_cast<size_t>(std::distance(schema.fields.begin(), field));
    };
    const size_t enabledIndex = indexOf(L"enabled");
    const size_t countIndex   = indexOf(L"count");
    const size_t textIndex    = indexOf(L"text");
    const size_t modeIndex    = indexOf(L"mode");
    const size_t tagsIndex    = indexOf(L"tags");
    Check(enabledIndex != kNoFieldIndex && ! configuration.values[enabledIndex].boolean,
          L"plugin config model: bool token coercion matches both former editors",
          success);
    Check(countIndex != kNoFieldIndex && configuration.values[countIndex].integer == 7,
          L"plugin config model: numeric values retain integral conversion behavior",
          success);
    Check(tagsIndex != kNoFieldIndex && configuration.values[tagsIndex].selection == std::vector<std::wstring>{L"a", L"future-option"},
          L"plugin config model: selection arrays skip wrong types without losing future string values",
          success);

    configuration.values[textIndex].text    = L"edited";
    configuration.values[countIndex].integer = 99;
    configuration.values[enabledIndex].boolean = true;
    configuration.values[modeIndex].text = L"a";
    configuration.values[tagsIndex].selection = {L"b"};
    std::string serialized;
    const HRESULT serializeHr = SerializeConfiguration(configurationJson, schema.fields, configuration.values, serialized);
    Check(serializeHr == S_OK && ! serialized.empty(), L"plugin config codec: edited values serialize successfully", success);

    Common::Settings::JsonValue serializedRoot;
    Check(Common::Settings::ParseJsonValue(serialized, serializedRoot) == S_OK,
          L"plugin config codec: serialized output is valid JSON",
          success);
    Check(Common::Settings::FindMember(serializedRoot, "futureBefore") != nullptr &&
              Common::Settings::FindMember(serializedRoot, "futureAfter") != nullptr,
          L"plugin config codec: unknown future members survive edits",
          success);
    Check(Common::Settings::GetWString(serializedRoot, "text") == std::make_optional<std::wstring>(L"edited") &&
              Common::Settings::GetBool(serializedRoot, "enabled") == std::make_optional(true),
          L"plugin config codec: text and bool edits are applied",
          success);
    const Common::Settings::JsonValue* countValue = Common::Settings::FindMember(serializedRoot, "count");
    const bool countIsNine = countValue &&
                             ((std::get_if<int64_t>(&countValue->value) && *std::get_if<int64_t>(&countValue->value) == 9) ||
                              (std::get_if<uint64_t>(&countValue->value) && *std::get_if<uint64_t>(&countValue->value) == 9u));
    Check(countIsNine,
          L"plugin config codec: integer constraints clamp during serialization",
          success);

    constexpr std::string_view invalidSchema = R"json({ "fields": [
        { "key": "duplicate", "type": "value", "min": 10, "max": 1 },
        { "key": "duplicate", "type": "text" },
        { "key": "future", "type": "future-type" }
    ] })json";
    const SchemaParseResult invalid = ParseSchema(invalidSchema);
    Check(invalid.HasErrors() && HasPluginConfigurationIssue(invalid, ValidationCode::InvalidConstraintRange) &&
              HasPluginConfigurationIssue(invalid, ValidationCode::DuplicateFieldKey) &&
              HasPluginConfigurationIssue(invalid, ValidationCode::UnsupportedType),
          L"plugin config schema: invalid constraints, duplicate IDs, and future types have explicit validation results",
          success);
}

} // namespace

int wmain(int argc, wchar_t** argv)
{
    bool success = true;

    std::wcout << L"[----------] SettingsSchemaParser tests\n";

    // Step 1 smoke: parser links and tolerates trivial input.
    Check(SettingsSchemaParser::ParseSettingsSchema("{}").empty(), L"smoke: ParseSettingsSchema(\"{}\") returns empty", success);

    // Step 3: synthetic-input tests (no external file needed).
    TestSyntheticDefaults(success);
    TestSyntheticMinMax(success);
    TestSyntheticEnum(success);
    TestSyntheticNoPane(success);
    TestMalformedInput(success);
    TestMissingFile(success);
    TestNonAscii(success);
    TestPluginConfigurationModelAndCodec(success);

    // Step 2: real-schema characterization (requires the schema file).
    const std::wstring schemaPath = FindSchemaPath(argc, argv);
    Check(! schemaPath.empty(), L"real schema: Specs\\SettingsStore.schema.json located", success);
    if (! schemaPath.empty())
    {
        std::wcout << L"[----------] using schema: " << schemaPath << L"\n";
        TestRealSchema(schemaPath, success);
    }

    if (success)
    {
        std::wcout << L"[  PASSED  ] all SettingsSchemaParser tests\n";
        return 0;
    }

    std::wcerr << L"[  FAILED  ] one or more SettingsSchemaParser tests\n";
    return 1;
}
