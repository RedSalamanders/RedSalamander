// Preferences.Plugin.Configuration.cpp

#include "Framework.h"

#include "Preferences.Plugin.Configuration.h"

#include <algorithm>
#include <array>
#include <cerrno>
#include <cwchar>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include <commctrl.h>
#include <commdlg.h>
#include <shobjidl.h>
#include <yyjson.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "FileSystemPluginManager.h"
#include "Helpers.h"
#include "ThemedControls.h"
#include "ViewerPluginManager.h"

#include "resource.h"

namespace
{
[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
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

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
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

[[nodiscard]] bool TryBrowseFolderPath(HWND owner, std::filesystem::path& outPath) noexcept
{
    outPath.clear();

    wil::com_ptr<IFileOpenDialog> dialog;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.put()));
    if (FAILED(hr) || ! dialog)
    {
        return false;
    }

    DWORD options = 0;
    if (FAILED(dialog->GetOptions(&options)))
    {
        return false;
    }
    static_cast<void>(dialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST));

    if (FAILED(dialog->Show(owner)))
    {
        return false;
    }

    wil::com_ptr<IShellItem> result;
    if (FAILED(dialog->GetResult(result.put())) || ! result)
    {
        return false;
    }

    wil::unique_cotaskmem_string selectedPath;
    if (FAILED(result->GetDisplayName(SIGDN_FILESYSPATH, selectedPath.put())) || ! selectedPath)
    {
        return false;
    }

    outPath = std::filesystem::path(selectedPath.get());
    return ! outPath.empty();
}

PrefsPluginConfigFieldType ParsePluginConfigFieldType(std::string_view type) noexcept
{
    if (type == "text")
    {
        return PrefsPluginConfigFieldType::Text;
    }
    if (type == "value")
    {
        return PrefsPluginConfigFieldType::Value;
    }
    if (type == "bool" || type == "boolean")
    {
        return PrefsPluginConfigFieldType::Bool;
    }
    if (type == "option")
    {
        return PrefsPluginConfigFieldType::Option;
    }
    if (type == "selection")
    {
        return PrefsPluginConfigFieldType::Selection;
    }
    return PrefsPluginConfigFieldType::Text;
}

std::optional<std::string_view> TryGetUtf8String(yyjson_val* obj, const char* key) noexcept
{
    if (! obj || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v || ! yyjson_is_str(v))
    {
        return std::nullopt;
    }

    const char* s = yyjson_get_str(v);
    if (! s)
    {
        return std::nullopt;
    }

    return std::string_view(s);
}

bool TryGetInt64(yyjson_val* obj, const char* key, int64_t& out) noexcept
{
    if (! obj || ! key)
    {
        return false;
    }

    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v)
    {
        return false;
    }

    if (yyjson_is_sint(v))
    {
        out = yyjson_get_sint(v);
        return true;
    }

    if (yyjson_is_uint(v))
    {
        out = static_cast<int64_t>(std::min<uint64_t>(yyjson_get_uint(v), static_cast<uint64_t>(std::numeric_limits<int64_t>::max())));
        return true;
    }

    if (yyjson_is_real(v))
    {
        out = static_cast<int64_t>(yyjson_get_real(v));
        return true;
    }

    return false;
}

[[nodiscard]] bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() > static_cast<size_t>(std::numeric_limits<int>::max()) || b.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    return CompareStringOrdinal(a.data(), static_cast<int>(a.size()), b.data(), static_cast<int>(b.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] std::optional<bool> TryParseBoolToggleToken(const std::wstring_view token) noexcept
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

bool TryGetBoolValue(yyjson_val* obj, const char* key, bool& out) noexcept
{
    if (! obj || ! key)
    {
        return false;
    }

    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v)
    {
        return false;
    }

    if (yyjson_is_bool(v))
    {
        out = yyjson_get_bool(v);
        return true;
    }

    if (yyjson_is_sint(v))
    {
        out = yyjson_get_sint(v) != 0;
        return true;
    }

    if (yyjson_is_uint(v))
    {
        out = yyjson_get_uint(v) != 0;
        return true;
    }

    if (yyjson_is_str(v))
    {
        const char* s = yyjson_get_str(v);
        if (! s)
        {
            return false;
        }

        const std::optional<bool> parsed = TryParseBoolToggleToken(Utf16FromUtf8(s));
        if (parsed.has_value())
        {
            out = parsed.value();
            return true;
        }
    }

    return false;
}

std::vector<PrefsPluginConfigField> ParsePluginConfigSchema(std::string_view schemaJsonUtf8) noexcept
{
    std::vector<PrefsPluginConfigField> fields;
    if (schemaJsonUtf8.empty())
    {
        return fields;
    }

    yyjson_doc* doc = yyjson_read(schemaJsonUtf8.data(), schemaJsonUtf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
    if (! doc)
    {
        return fields;
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return fields;
    }

    yyjson_val* fieldsArr = yyjson_obj_get(root, "fields");
    if (! fieldsArr || ! yyjson_is_arr(fieldsArr))
    {
        return fields;
    }

    const size_t count = yyjson_arr_size(fieldsArr);
    fields.reserve(count);

    for (size_t i = 0; i < count; ++i)
    {
        yyjson_val* item = yyjson_arr_get(fieldsArr, i);
        if (! item || ! yyjson_is_obj(item))
        {
            continue;
        }

        const auto keyUtf8  = TryGetUtf8String(item, "key");
        const auto typeUtf8 = TryGetUtf8String(item, "type");
        if (! keyUtf8.has_value() || ! typeUtf8.has_value())
        {
            continue;
        }

        PrefsPluginConfigField field;
        field.key = Utf16FromUtf8(keyUtf8.value());
        if (field.key.empty())
        {
            continue;
        }

        field.type = ParsePluginConfigFieldType(typeUtf8.value());

        const auto labelUtf8 = TryGetUtf8String(item, "label");
        field.label          = labelUtf8.has_value() ? Utf16FromUtf8(labelUtf8.value()) : field.key;
        if (field.label.empty())
        {
            field.label = field.key;
        }

        const auto descriptionUtf8 = TryGetUtf8String(item, "description");
        if (descriptionUtf8.has_value())
        {
            field.description = Utf16FromUtf8(descriptionUtf8.value());
        }

        if (field.type == PrefsPluginConfigFieldType::Text)
        {
            const auto browseUtf8 = TryGetUtf8String(item, "browse");
            if (browseUtf8.has_value())
            {
                const std::string_view browse = browseUtf8.value();
                field.browseFolder            = (browse == "folder") || (browse == "directory");
            }
        }

        int64_t minValue = 0;
        if (TryGetInt64(item, "min", minValue))
        {
            field.hasMin = true;
            field.min    = minValue;
        }

        int64_t maxValue = 0;
        if (TryGetInt64(item, "max", maxValue))
        {
            field.hasMax = true;
            field.max    = maxValue;
        }

        if (field.type == PrefsPluginConfigFieldType::Text)
        {
            const auto def    = TryGetUtf8String(item, "default");
            field.defaultText = def.has_value() ? Utf16FromUtf8(def.value()) : std::wstring();
        }
        else if (field.type == PrefsPluginConfigFieldType::Value)
        {
            int64_t defValue = 0;
            if (TryGetInt64(item, "default", defValue))
            {
                field.defaultInt = defValue;
            }
        }
        else if (field.type == PrefsPluginConfigFieldType::Bool)
        {
            bool def = false;
            if (TryGetBoolValue(item, "default", def))
            {
                field.defaultBool = def;
            }
        }
        else if (field.type == PrefsPluginConfigFieldType::Option)
        {
            const auto def      = TryGetUtf8String(item, "default");
            field.defaultOption = def.has_value() ? Utf16FromUtf8(def.value()) : std::wstring();

            yyjson_val* options = yyjson_obj_get(item, "options");
            if (options && yyjson_is_arr(options))
            {
                const size_t optCount = yyjson_arr_size(options);
                field.choices.reserve(optCount);
                for (size_t o = 0; o < optCount; ++o)
                {
                    yyjson_val* opt = yyjson_arr_get(options, o);
                    if (! opt || ! yyjson_is_obj(opt))
                    {
                        continue;
                    }

                    const auto valueUtf8 = TryGetUtf8String(opt, "value");
                    if (! valueUtf8.has_value())
                    {
                        continue;
                    }

                    PrefsPluginConfigChoice choice;
                    choice.value = Utf16FromUtf8(valueUtf8.value());
                    if (choice.value.empty())
                    {
                        continue;
                    }

                    const auto optLabelUtf8 = TryGetUtf8String(opt, "label");
                    choice.label            = optLabelUtf8.has_value() ? Utf16FromUtf8(optLabelUtf8.value()) : choice.value;
                    if (choice.label.empty())
                    {
                        choice.label = choice.value;
                    }

                    field.choices.push_back(std::move(choice));
                }
            }
        }
        else if (field.type == PrefsPluginConfigFieldType::Selection)
        {
            yyjson_val* options = yyjson_obj_get(item, "options");
            if (options && yyjson_is_arr(options))
            {
                const size_t optCount = yyjson_arr_size(options);
                field.choices.reserve(optCount);
                for (size_t o = 0; o < optCount; ++o)
                {
                    yyjson_val* opt = yyjson_arr_get(options, o);
                    if (! opt || ! yyjson_is_obj(opt))
                    {
                        continue;
                    }

                    const auto valueUtf8 = TryGetUtf8String(opt, "value");
                    if (! valueUtf8.has_value())
                    {
                        continue;
                    }

                    PrefsPluginConfigChoice choice;
                    choice.value = Utf16FromUtf8(valueUtf8.value());
                    if (choice.value.empty())
                    {
                        continue;
                    }

                    const auto optLabelUtf8 = TryGetUtf8String(opt, "label");
                    choice.label            = optLabelUtf8.has_value() ? Utf16FromUtf8(optLabelUtf8.value()) : choice.value;
                    if (choice.label.empty())
                    {
                        choice.label = choice.value;
                    }

                    field.choices.push_back(std::move(choice));
                }
            }

            yyjson_val* def = yyjson_obj_get(item, "default");
            if (def && yyjson_is_arr(def))
            {
                const size_t defCount = yyjson_arr_size(def);
                field.defaultSelection.reserve(defCount);
                for (size_t d = 0; d < defCount; ++d)
                {
                    yyjson_val* v = yyjson_arr_get(def, d);
                    if (! v || ! yyjson_is_str(v))
                    {
                        continue;
                    }

                    const char* s = yyjson_get_str(v);
                    if (! s)
                    {
                        continue;
                    }

                    std::wstring value = Utf16FromUtf8(s);
                    if (! value.empty())
                    {
                        field.defaultSelection.push_back(std::move(value));
                    }
                }
            }
        }

        fields.push_back(std::move(field));
    }

    return fields;
}

yyjson_doc* ParseJsonToDoc(std::string_view textUtf8) noexcept
{
    if (textUtf8.empty())
    {
        return nullptr;
    }

    return yyjson_read(textUtf8.data(), textUtf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
}

void ApplyFieldDefaultToControls(const PrefsPluginConfigField& field, PrefsPluginConfigFieldControls& out, yyjson_val* configRoot) noexcept
{
    out.field = field;

    const std::string keyUtf8 = Utf8FromUtf16(field.key);
    yyjson_val* current       = nullptr;
    if (configRoot && ! keyUtf8.empty())
    {
        current = yyjson_obj_get(configRoot, keyUtf8.c_str());
    }

    if (field.type == PrefsPluginConfigFieldType::Text)
    {
        std::wstring value = field.defaultText;
        if (current && yyjson_is_str(current))
        {
            const char* s = yyjson_get_str(current);
            if (s)
            {
                value = Utf16FromUtf8(s);
            }
        }
        out.field.defaultText = value;
    }
    else if (field.type == PrefsPluginConfigFieldType::Value)
    {
        int64_t value = field.defaultInt;
        if (current)
        {
            if (yyjson_is_sint(current))
            {
                value = yyjson_get_sint(current);
            }
            else if (yyjson_is_uint(current))
            {
                value = static_cast<int64_t>(std::min<uint64_t>(yyjson_get_uint(current), static_cast<uint64_t>(std::numeric_limits<int64_t>::max())));
            }
            else if (yyjson_is_real(current))
            {
                value = static_cast<int64_t>(yyjson_get_real(current));
            }
        }
        out.field.defaultInt = value;
    }
    else if (field.type == PrefsPluginConfigFieldType::Bool)
    {
        bool value = field.defaultBool;
        if (current)
        {
            if (yyjson_is_bool(current))
            {
                value = yyjson_get_bool(current);
            }
            else if (yyjson_is_sint(current))
            {
                value = yyjson_get_sint(current) != 0;
            }
            else if (yyjson_is_uint(current))
            {
                value = yyjson_get_uint(current) != 0;
            }
            else if (yyjson_is_str(current))
            {
                const char* s = yyjson_get_str(current);
                if (s)
                {
                    const std::optional<bool> parsed = TryParseBoolToggleToken(Utf16FromUtf8(s));
                    if (parsed.has_value())
                    {
                        value = parsed.value();
                    }
                }
            }
        }
        out.field.defaultBool = value;
    }
    else if (field.type == PrefsPluginConfigFieldType::Option)
    {
        out.schemaDefaultOption = field.defaultOption;
        std::wstring value      = field.defaultOption;
        if (current && yyjson_is_str(current))
        {
            const char* s = yyjson_get_str(current);
            if (s)
            {
                value = Utf16FromUtf8(s);
            }
        }
        out.field.defaultOption = value;
    }
    else if (field.type == PrefsPluginConfigFieldType::Selection)
    {
        std::vector<std::wstring> values = field.defaultSelection;
        if (current && yyjson_is_arr(current))
        {
            values.clear();
            const size_t count = yyjson_arr_size(current);
            values.reserve(count);
            for (size_t i = 0; i < count; ++i)
            {
                yyjson_val* v = yyjson_arr_get(current, i);
                if (! v || ! yyjson_is_str(v))
                {
                    continue;
                }
                const char* s = yyjson_get_str(v);
                if (! s)
                {
                    continue;
                }
                std::wstring t = Utf16FromUtf8(s);
                if (! t.empty())
                {
                    values.push_back(std::move(t));
                }
            }
        }

        out.field.defaultSelection = std::move(values);
    }
}
} // namespace

namespace
{
[[nodiscard]] bool ContainsChoiceValue(const std::vector<std::wstring>& values, std::wstring_view needle) noexcept;
[[nodiscard]] size_t FindChoiceIndex(const std::vector<PrefsPluginConfigChoice>& choices, std::wstring_view desired) noexcept;
[[nodiscard]] std::wstring GetPluginConfigurationSchemaErrorText(const PrefsPluginListItem& pluginItem) noexcept;
[[nodiscard]] PrefsPluginConfigFieldControls* FindFieldForControl(PreferencesDialogState& state, HWND hwnd) noexcept;
[[nodiscard]] bool CommitEditor(HWND host, PreferencesDialogState& state) noexcept;
} // namespace

namespace PrefsPluginConfiguration
{
void Clear(PreferencesDialogState& state) noexcept
{
    state.pluginsDetailsConfigFields.clear();
    state.pluginsDetailsConfigPluginId.clear();
}

[[nodiscard]] bool EnsureEditor(HWND parent, PreferencesDialogState& state, const PrefsPluginListItem& pluginItem) noexcept
{
    if (! parent)
    {
        return false;
    }

    const std::wstring_view pluginId = PrefsPlugins::GetId(pluginItem);
    if (pluginId.empty())
    {
        Clear(state);
        return false;
    }

    if (state.pluginsDetailsConfigPluginId == pluginId && ! state.pluginsDetailsConfigFields.empty())
    {
        const auto isValid = [](HWND hwnd) noexcept { return ! hwnd || IsWindow(hwnd); };
        bool valid         = true;

        for (const PrefsPluginConfigFieldControls& controls : state.pluginsDetailsConfigFields)
        {
            valid = valid && isValid(controls.label.get());
            valid = valid && isValid(controls.description.get());
            valid = valid && isValid(controls.editFrame.get());
            valid = valid && isValid(controls.edit.get());
            valid = valid && isValid(controls.browseButton.get());
            valid = valid && isValid(controls.comboFrame.get());
            valid = valid && isValid(controls.combo.get());
            valid = valid && isValid(controls.toggle.get());
            for (const auto& button : controls.choiceButtons)
            {
                valid = valid && isValid(button.get());
            }
            if (! valid)
            {
                break;
            }
        }

        if (valid)
        {
            return true;
        }
    }

    Clear(state);
    state.pluginsDetailsConfigPluginId = std::wstring(pluginId);

    std::string schemaUtf8;
    HRESULT schemaHr = E_FAIL;
    if (pluginItem.type == PrefsPluginType::FileSystem)
    {
        schemaHr = FileSystemPluginManager::GetInstance().GetConfigurationSchema(pluginId, state.baselineSettings, schemaUtf8);
    }
    else
    {
        schemaHr = ViewerPluginManager::GetInstance().GetConfigurationSchema(pluginId, state.baselineSettings, schemaUtf8);
    }

    if (FAILED(schemaHr))
    {
        if (state.pluginsDetailsConfigError)
        {
            const std::wstring message = GetPluginConfigurationSchemaErrorText(pluginItem);
            SetWindowTextW(state.pluginsDetailsConfigError.get(), message.c_str());
        }
        return false;
    }

    const std::vector<PrefsPluginConfigField> fields = ParsePluginConfigSchema(schemaUtf8);
    if (fields.empty())
    {
        if (state.pluginsDetailsConfigError)
        {
            const std::wstring message = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_SCHEMA_NO_FIELDS);
            SetWindowTextW(state.pluginsDetailsConfigError.get(), message.c_str());
        }
        return false;
    }

    std::string configUtf8;
    const std::wstring pluginIdText(pluginId);
    const auto it = state.workingSettings.plugins.configurationByPluginId.find(pluginIdText);
    if (it != state.workingSettings.plugins.configurationByPluginId.end() && ! std::holds_alternative<std::monostate>(it->second.value))
    {
        static_cast<void>(Common::Settings::SerializeJsonValue(it->second, configUtf8));
    }

    if (configUtf8.empty())
    {
        HRESULT configHr = E_FAIL;
        if (pluginItem.type == PrefsPluginType::FileSystem)
        {
            configHr = FileSystemPluginManager::GetInstance().GetConfiguration(pluginId, state.baselineSettings, configUtf8);
        }
        else
        {
            configHr = ViewerPluginManager::GetInstance().GetConfiguration(pluginId, state.baselineSettings, configUtf8);
        }

        if (FAILED(configHr))
        {
            configUtf8.clear();
        }
    }

    if (configUtf8.empty())
    {
        configUtf8 = "{}";
    }

    yyjson_doc* configDoc = ParseJsonToDoc(configUtf8);
    auto freeConfigDoc    = wil::scope_exit(
        [&]
        {
            if (configDoc)
            {
                yyjson_doc_free(configDoc);
            }
        });

    yyjson_val* configRoot = nullptr;
    if (configDoc)
    {
        configRoot = yyjson_doc_get_root(configDoc);
        if (! configRoot || ! yyjson_is_obj(configRoot))
        {
            configRoot = nullptr;
        }
    }

    if (state.pluginsDetailsConfigError)
    {
        SetWindowTextW(state.pluginsDetailsConfigError.get(), L"");
    }

    const HWND panel = parent;

    const DWORD baseStaticStyle   = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    const DWORD wrapStaticStyle   = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;
    const DWORD browseButtonStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON | (state.theme.systemHighContrast ? 0U : BS_OWNERDRAW);

    state.pluginsDetailsConfigFields.clear();
    state.pluginsDetailsConfigFields.reserve(fields.size());

    for (const PrefsPluginConfigField& field : fields)
    {
        PrefsPluginConfigFieldControls controls{};
        ApplyFieldDefaultToControls(field, controls, configRoot);

        controls.label.reset(
            CreateWindowExW(0, L"Static", controls.field.label.c_str(), baseStaticStyle, 0, 0, 10, 10, panel, nullptr, GetModuleHandleW(nullptr), nullptr));

        controls.description.reset(CreateWindowExW(
            0, L"Static", controls.field.description.c_str(), wrapStaticStyle, 0, 0, 10, 10, panel, nullptr, GetModuleHandleW(nullptr), nullptr));

        if (controls.field.type == PrefsPluginConfigFieldType::Text || controls.field.type == PrefsPluginConfigFieldType::Value)
        {
            DWORD editStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL;
            if (controls.field.type == PrefsPluginConfigFieldType::Value)
            {
                editStyle |= ES_NUMBER;
            }

            PrefsInput::CreateFramedEditBox(state, panel, controls.editFrame, controls.edit, 0, editStyle);
            if (controls.edit)
            {
                if (controls.field.type == PrefsPluginConfigFieldType::Text)
                {
                    SetWindowTextW(controls.edit.get(), controls.field.defaultText.c_str());
                }
                else
                {
                    const std::wstring text = std::to_wstring(controls.field.defaultInt);
                    SetWindowTextW(controls.edit.get(), text.c_str());
                }
            }

            if (controls.field.type == PrefsPluginConfigFieldType::Text && controls.field.browseFolder)
            {
                const std::wstring label = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_CONFIG_BROWSE_ELLIPSIS);
                controls.browseButton.reset(
                    CreateWindowExW(0, L"Button", label.c_str(), browseButtonStyle, 0, 0, 10, 10, panel, nullptr, GetModuleHandleW(nullptr), nullptr));
                if (controls.browseButton)
                {
                    PrefsInput::EnableMouseWheelForwarding(controls.browseButton.get());
                }
            }
        }
        else if (controls.field.type == PrefsPluginConfigFieldType::Bool)
        {
            if (! state.theme.systemHighContrast)
            {
                controls.toggle.reset(CreateWindowExW(
                    0, L"Button", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW, 0, 0, 10, 10, panel, nullptr, GetModuleHandleW(nullptr), nullptr));
                if (controls.toggle)
                {
                    SetWindowLongPtrW(controls.toggle.get(), GWLP_USERDATA, controls.field.defaultBool ? 1 : 0);
                    PrefsInput::EnableMouseWheelForwarding(controls.toggle.get());
                }
            }
            else
            {
                controls.toggle.reset(CreateWindowExW(
                    0, L"Button", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 10, 10, panel, nullptr, GetModuleHandleW(nullptr), nullptr));
                if (controls.toggle)
                {
                    SendMessageW(controls.toggle.get(), BM_SETCHECK, controls.field.defaultBool ? BST_CHECKED : BST_UNCHECKED, 0);
                    PrefsInput::EnableMouseWheelForwarding(controls.toggle.get());
                }
            }
        }
        else if (controls.field.type == PrefsPluginConfigFieldType::Option)
        {
            if (! state.theme.systemHighContrast && controls.field.choices.size() == 2)
            {
                const size_t notFound = static_cast<size_t>(std::numeric_limits<size_t>::max());
                size_t defaultIndex   = FindChoiceIndex(controls.field.choices, controls.schemaDefaultOption);
                if (defaultIndex == notFound)
                {
                    defaultIndex = 0;
                }

                const size_t otherIndex       = defaultIndex == 0 ? 1 : 0;
                controls.toggleOnChoiceIndex  = defaultIndex;
                controls.toggleOffChoiceIndex = otherIndex;

                const size_t currentIndex = FindChoiceIndex(controls.field.choices, controls.field.defaultOption);
                const bool toggledOn      = (currentIndex == defaultIndex) || (currentIndex == notFound);

                controls.toggle.reset(CreateWindowExW(
                    0, L"Button", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW, 0, 0, 10, 10, panel, nullptr, GetModuleHandleW(nullptr), nullptr));
                if (controls.toggle)
                {
                    SetWindowLongPtrW(controls.toggle.get(), GWLP_USERDATA, toggledOn ? 1 : 0);
                    PrefsInput::EnableMouseWheelForwarding(controls.toggle.get());
                }
            }
            else
            {
                PrefsInput::CreateFramedComboBox(state, panel, controls.comboFrame, controls.combo, 0);
                if (controls.combo)
                {
                    SendMessageW(controls.combo.get(), CB_RESETCONTENT, 0, 0);
                    for (const auto& choice : controls.field.choices)
                    {
                        const std::wstring_view label = choice.label.empty() ? std::wstring_view(choice.value) : std::wstring_view(choice.label);
                        SendMessageW(controls.combo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.data()));
                    }

                    size_t selectedIndex = FindChoiceIndex(controls.field.choices, controls.field.defaultOption);
                    if (selectedIndex == static_cast<size_t>(std::numeric_limits<size_t>::max()))
                    {
                        selectedIndex = 0;
                    }

                    SendMessageW(controls.combo.get(), CB_SETCURSEL, static_cast<WPARAM>(selectedIndex), 0);
                    PrefsUi::InvalidateComboBox(controls.combo.get());
                    ThemedControls::ApplyThemeToComboBox(controls.combo.get(), state.theme);
                    ThemedControls::EnsureComboBoxDroppedWidth(controls.combo.get(), GetDpiForWindow(controls.combo.get()));
                }
            }
        }
        else if (controls.field.type == PrefsPluginConfigFieldType::Selection)
        {
            const std::vector<std::wstring>& selected = controls.field.defaultSelection;
            controls.choiceButtons.reserve(controls.field.choices.size());
            for (const auto& choice : controls.field.choices)
            {
                const std::wstring_view label = choice.label.empty() ? std::wstring_view(choice.value) : std::wstring_view(choice.label);
                HWND button                   = CreateWindowExW(0,
                                              L"Button",
                                              label.data(),
                                              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                                              0,
                                              0,
                                              10,
                                              10,
                                              panel,
                                              nullptr,
                                              GetModuleHandleW(nullptr),
                                              nullptr);
                if (button)
                {
                    const bool checked = ContainsChoiceValue(selected, choice.value);
                    SendMessageW(button, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
                    PrefsInput::EnableMouseWheelForwarding(button);
                }
                controls.choiceButtons.emplace_back(button);
            }
        }

        state.pluginsDetailsConfigFields.push_back(std::move(controls));
    }

    return ! state.pluginsDetailsConfigFields.empty();
}

void LayoutCards(HWND host, PreferencesDialogState& state, int x, int& y, int width, HFONT dialogFont) noexcept
{
    if (! host || width <= 0)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(host);

    const int rowHeight      = std::max(1, ThemedControls::ScaleDip(dpi, 26));
    const int titleHeight    = std::max(1, ThemedControls::ScaleDip(dpi, 18));
    const int optionHeight   = std::max(1, ThemedControls::ScaleDip(dpi, 20));
    const int minToggleWidth = ThemedControls::ScaleDip(dpi, 90);

    const int cardPaddingX = ThemedControls::ScaleDip(dpi, 12);
    const int cardPaddingY = ThemedControls::ScaleDip(dpi, 8);
    const int cardGapY     = ThemedControls::ScaleDip(dpi, 2);
    const int cardGapX     = ThemedControls::ScaleDip(dpi, 12);
    const int cardSpacingY = ThemedControls::ScaleDip(dpi, 8);
    const int innerGapX    = ThemedControls::ScaleDip(dpi, 8);
    const int buttonPadX   = ThemedControls::ScaleDip(dpi, 12);

    const int maxControlWidth = std::max(0, width - 2 * cardPaddingX);

    const HFONT infoFont = state.italicFont ? state.italicFont.get() : dialogFont;

    const auto pushCard = [&](const RECT& card) noexcept { state.pageSettingCards.push_back(card); };

    const auto measureButtonWidth = [&](HWND button, int minWidthDip) noexcept
    {
        if (! button)
        {
            return 0;
        }

        HFONT font = reinterpret_cast<HFONT>(SendMessageW(button, WM_GETFONT, 0, 0));
        if (! font)
        {
            font = dialogFont ? dialogFont : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        }

        const std::wstring text = PrefsUi::GetWindowTextString(button);
        const int textW         = ThemedControls::MeasureTextWidth(host, font, text);
        return std::max(ThemedControls::ScaleDip(dpi, minWidthDip), textW + 2 * buttonPadX);
    };

    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

    for (PrefsPluginConfigFieldControls& controls : state.pluginsDetailsConfigFields)
    {
        const bool isSelection = controls.field.type == PrefsPluginConfigFieldType::Selection;

        const std::wstring descText = controls.description ? PrefsUi::GetWindowTextString(controls.description.get()) : std::wstring{};
        const bool hasDesc          = ! descText.empty();

        int controlGroupWidth = 0;
        int editWidth         = 0;
        int browseWidth       = 0;

        if (! isSelection)
        {
            if (controls.edit)
            {
                const int minEditWidth = ThemedControls::ScaleDip(dpi, 140);
                int desiredWidth       = minEditWidth;

                if (controls.field.type == PrefsPluginConfigFieldType::Text)
                {
                    desiredWidth = ThemedControls::ScaleDip(dpi, controls.field.browseFolder ? 380 : 320);
                }

                if (controls.field.type == PrefsPluginConfigFieldType::Value)
                {
                    desiredWidth = ThemedControls::ScaleDip(dpi, 140);
                }

                browseWidth = controls.browseButton ? measureButtonWidth(controls.browseButton.get(), 90) : 0;
                if (browseWidth > 0)
                {
                    const int maxBrowseWidth = std::max(0, maxControlWidth - innerGapX - 1);
                    browseWidth              = std::min(browseWidth, maxBrowseWidth);
                }

                const int browseExtra  = (browseWidth > 0) ? (innerGapX + browseWidth) : 0;
                const int maxEditWidth = std::max(1, maxControlWidth - browseExtra);

                editWidth = std::clamp(desiredWidth, 1, maxEditWidth);
                if (maxEditWidth >= minEditWidth)
                {
                    editWidth = std::max(minEditWidth, editWidth);
                }
                controlGroupWidth = editWidth + browseExtra;
            }
            else if (controls.combo)
            {
                int desiredWidth  = ThemedControls::MeasureComboBoxPreferredWidth(controls.combo.get(), dpi);
                desiredWidth      = std::max(desiredWidth, ThemedControls::ScaleDip(dpi, 160));
                desiredWidth      = std::min(desiredWidth, std::min(maxControlWidth, ThemedControls::ScaleDip(dpi, 260)));
                controlGroupWidth = desiredWidth;
            }
            else if (controls.toggle)
            {
                int desiredWidth = std::min(maxControlWidth, ThemedControls::ScaleDip(dpi, 180));
                if (! state.theme.systemHighContrast)
                {
                    std::wstring_view onStateLabel  = onLabel;
                    std::wstring_view offStateLabel = offLabel;

                    if (controls.field.type == PrefsPluginConfigFieldType::Option && controls.field.choices.size() >= 2)
                    {
                        const auto& choices   = controls.field.choices;
                        const size_t onIndex  = std::min(controls.toggleOnChoiceIndex, choices.size() - 1);
                        const size_t offIndex = std::min(controls.toggleOffChoiceIndex, choices.size() - 1);

                        onStateLabel = choices[onIndex].label.empty() ? std::wstring_view(choices[onIndex].value) : std::wstring_view(choices[onIndex].label);
                        offStateLabel =
                            choices[offIndex].label.empty() ? std::wstring_view(choices[offIndex].value) : std::wstring_view(choices[offIndex].label);
                    }

                    const HFONT toggleFont = state.boldFont ? state.boldFont.get() : dialogFont;
                    const int paddingX     = ThemedControls::ScaleDip(dpi, 6);
                    const int gapX         = ThemedControls::ScaleDip(dpi, 8);
                    const int trackWidth   = ThemedControls::ScaleDip(dpi, 34);

                    const int onWidth        = ThemedControls::MeasureTextWidth(host, toggleFont, onStateLabel);
                    const int offWidth       = ThemedControls::MeasureTextWidth(host, toggleFont, offStateLabel);
                    const int stateTextWidth = std::max(onWidth, offWidth);
                    const int measured       = std::max(minToggleWidth, (2 * paddingX) + stateTextWidth + gapX + trackWidth);
                    desiredWidth             = std::min(maxControlWidth, measured);
                }
                else
                {
                    desiredWidth = std::min(maxControlWidth, rowHeight);
                }

                controlGroupWidth = desiredWidth;
            }
        }

        const int textWidth = std::max(0, width - 2 * cardPaddingX - ((controlGroupWidth > 0) ? (cardGapX + controlGroupWidth) : 0));

        int descHeight = 0;
        if (hasDesc && controls.description)
        {
            descHeight = PrefsUi::MeasureStaticTextHeight(host, infoFont, textWidth, descText);
        }

        int cardHeight = 0;
        if (isSelection)
        {
            const int optionCount   = static_cast<int>(controls.choiceButtons.size());
            const int optionsHeight = std::max(0, optionCount * optionHeight);

            int contentHeight = titleHeight;
            if (optionsHeight > 0)
            {
                contentHeight += cardGapY + optionsHeight;
            }
            if (hasDesc)
            {
                contentHeight += cardGapY + descHeight;
            }
            cardHeight = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);
        }
        else
        {
            const int contentHeight = hasDesc ? (titleHeight + cardGapY + descHeight) : titleHeight;
            cardHeight              = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);
        }

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        pushCard(card);

        const int controlX = card.right - cardPaddingX - controlGroupWidth;
        const int controlY = card.top + cardPaddingY;

        if (controls.label)
        {
            SetWindowPos(
                controls.label.get(), nullptr, card.left + cardPaddingX, card.top + cardPaddingY, textWidth, titleHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(controls.label.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        if (! isSelection)
        {
            if (controls.editFrame && controls.edit)
            {
                const int framePadding = ThemedControls::ScaleDip(dpi, 1);
                SetWindowPos(controls.editFrame.get(), nullptr, controlX, controlY, editWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SetWindowPos(controls.edit.get(),
                             nullptr,
                             controlX + framePadding,
                             controlY + framePadding,
                             std::max(1, editWidth - 2 * framePadding),
                             std::max(1, rowHeight - 2 * framePadding),
                             SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(controls.edit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            if (controls.browseButton)
            {
                SetWindowPos(
                    controls.browseButton.get(), nullptr, controlX + editWidth + innerGapX, controlY, browseWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(controls.browseButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            if (controls.comboFrame && controls.combo)
            {
                const int framePadding = ThemedControls::ScaleDip(dpi, 1);
                SetWindowPos(controls.comboFrame.get(), nullptr, controlX, controlY, controlGroupWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SetWindowPos(controls.combo.get(),
                             nullptr,
                             controlX + framePadding,
                             controlY + framePadding,
                             std::max(1, controlGroupWidth - 2 * framePadding),
                             std::max(1, rowHeight - 2 * framePadding),
                             SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(controls.combo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
                ThemedControls::EnsureComboBoxDroppedWidth(controls.combo.get(), dpi);
            }
            else if (controls.toggle)
            {
                SetWindowPos(controls.toggle.get(), nullptr, controlX, controlY, controlGroupWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(controls.toggle.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }
        }
        else
        {
            int contentY = card.top + cardPaddingY + titleHeight;
            if (! controls.choiceButtons.empty())
            {
                contentY += cardGapY;
            }

            const int optionWidth = std::max(0, width - 2 * cardPaddingX);
            for (size_t i = 0; i < controls.choiceButtons.size(); ++i)
            {
                HWND button = controls.choiceButtons[i].get();
                if (! button)
                {
                    continue;
                }

                const int buttonY = contentY + static_cast<int>(i) * optionHeight;
                SetWindowPos(button, nullptr, card.left + cardPaddingX, buttonY, optionWidth, optionHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(button, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            contentY += static_cast<int>(controls.choiceButtons.size()) * optionHeight;
            if (hasDesc)
            {
                contentY += cardGapY;
            }
        }

        if (controls.description)
        {
            if (hasDesc)
            {
                int descY = 0;
                if (isSelection)
                {
                    descY = card.top + cardPaddingY + titleHeight;
                    if (! controls.choiceButtons.empty())
                    {
                        descY += cardGapY + static_cast<int>(controls.choiceButtons.size()) * optionHeight;
                    }
                    descY += cardGapY;
                }
                else
                {
                    descY = card.top + cardPaddingY + titleHeight + cardGapY;
                }

                SetWindowPos(
                    controls.description.get(), nullptr, card.left + cardPaddingX, descY, textWidth, std::max(0, descHeight), SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(controls.description.get(), WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
                ShowWindow(controls.description.get(), SW_SHOW);
            }
            else
            {
                ShowWindow(controls.description.get(), SW_HIDE);
            }
        }

        y += cardHeight + cardSpacingY;
    }
}

[[nodiscard]] bool HandleCommand(HWND host, PreferencesDialogState& state, UINT notifyCode, HWND hwndCtl) noexcept
{
    if (! host || ! hwndCtl || state.pluginsDetailsConfigFields.empty() || state.pluginsDetailsConfigPluginId.empty())
    {
        return false;
    }

    PrefsPluginConfigFieldControls* controls = FindFieldForControl(state, hwndCtl);
    if (! controls)
    {
        return false;
    }

    if (notifyCode == BN_CLICKED)
    {
        if (controls->browseButton.get() == hwndCtl && controls->edit && controls->field.browseFolder)
        {
            std::filesystem::path selectedPath;
            HWND owner = GetParent(host);
            if (! TryBrowseFolderPath(owner ? owner : host, selectedPath))
            {
                return true;
            }

            const std::wstring text = selectedPath.wstring();
            SetWindowTextW(controls->edit.get(), text.c_str());
            return CommitEditor(host, state);
        }

        if (controls->toggle.get() == hwndCtl)
        {
            const LONG_PTR style = GetWindowLongPtrW(hwndCtl, GWL_STYLE);
            if ((style & BS_TYPEMASK) == BS_OWNERDRAW)
            {
                const LONG_PTR current = GetWindowLongPtrW(hwndCtl, GWLP_USERDATA);
                SetWindowLongPtrW(hwndCtl, GWLP_USERDATA, current == 0 ? 1 : 0);
                InvalidateRect(hwndCtl, nullptr, TRUE);
            }
        }

        return CommitEditor(host, state);
    }

    if (notifyCode == EN_KILLFOCUS && controls->edit.get() == hwndCtl)
    {
        return CommitEditor(host, state);
    }

    if (notifyCode == CBN_SELCHANGE && controls->combo.get() == hwndCtl)
    {
        return CommitEditor(host, state);
    }

    return false;
}
} // namespace PrefsPluginConfiguration

namespace
{
std::string BuildConfigurationJson(const std::vector<PrefsPluginConfigFieldControls>& controls) noexcept
{
    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return {};
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    if (! root)
    {
        return {};
    }
    yyjson_mut_doc_set_root(doc, root);

    for (const auto& c : controls)
    {
        const std::string keyUtf8 = Utf8FromUtf16(c.field.key);
        if (keyUtf8.empty() && ! c.field.key.empty())
        {
            continue;
        }

        if (keyUtf8.empty())
        {
            continue;
        }

        yyjson_mut_val* key = yyjson_mut_strncpy(doc, keyUtf8.c_str(), keyUtf8.size());
        if (! key)
        {
            return {};
        }

        if (c.field.type == PrefsPluginConfigFieldType::Text)
        {
            std::wstring value;
            if (c.edit)
            {
                const int len = GetWindowTextLengthW(c.edit.get());
                if (len > 0)
                {
                    value.resize(static_cast<size_t>(len) + 1u);
                    GetWindowTextW(c.edit.get(), value.data(), len + 1);
                    value.resize(static_cast<size_t>(len));
                }
            }

            const std::string utf8 = Utf8FromUtf16(value);
            yyjson_mut_val* val    = yyjson_mut_strncpy(doc, utf8.c_str(), utf8.size());
            if (! val)
            {
                return {};
            }
            if (! yyjson_mut_obj_add(root, key, val))
            {
                return {};
            }
        }
        else if (c.field.type == PrefsPluginConfigFieldType::Value)
        {
            int64_t v = c.field.defaultInt;
            if (c.edit)
            {
                std::array<wchar_t, 64> buffer{};
                GetWindowTextW(c.edit.get(), buffer.data(), static_cast<int>(buffer.size()));

                wchar_t* end           = nullptr;
                errno                  = 0;
                const long long parsed = std::wcstoll(buffer.data(), &end, 10);
                if (errno == 0 && end != buffer.data())
                {
                    v = static_cast<int64_t>(parsed);
                }
            }

            if (c.field.hasMin)
            {
                v = std::max(v, c.field.min);
            }
            if (c.field.hasMax)
            {
                v = std::min(v, c.field.max);
            }

            yyjson_mut_val* val = yyjson_mut_int(doc, v);
            if (! val)
            {
                return {};
            }
            if (! yyjson_mut_obj_add(root, key, val))
            {
                return {};
            }
        }
        else if (c.field.type == PrefsPluginConfigFieldType::Bool)
        {
            bool v = c.field.defaultBool;
            if (c.toggle)
            {
                const LONG_PTR style = GetWindowLongPtrW(c.toggle.get(), GWL_STYLE);
                const LONG_PTR type  = style & BS_TYPEMASK;
                if (type == BS_OWNERDRAW)
                {
                    v = GetWindowLongPtrW(c.toggle.get(), GWLP_USERDATA) != 0;
                }
                else
                {
                    v = SendMessageW(c.toggle.get(), BM_GETCHECK, 0, 0) == BST_CHECKED;
                }
            }
            else if (! c.choiceButtons.empty())
            {
                v = SendMessageW(c.choiceButtons.front().get(), BM_GETCHECK, 0, 0) == BST_CHECKED;
            }

            yyjson_mut_val* val = yyjson_mut_bool(doc, v ? true : false);
            if (! val)
            {
                return {};
            }
            if (! yyjson_mut_obj_add(root, key, val))
            {
                return {};
            }
        }
        else if (c.field.type == PrefsPluginConfigFieldType::Option)
        {
            std::wstring selected;
            if (c.toggle)
            {
                const bool isOn    = GetWindowLongPtrW(c.toggle.get(), GWLP_USERDATA) != 0;
                const size_t index = isOn ? c.toggleOnChoiceIndex : c.toggleOffChoiceIndex;
                if (index < c.field.choices.size())
                {
                    selected = c.field.choices[index].value;
                }
            }
            else if (c.combo)
            {
                const LRESULT index = SendMessageW(c.combo.get(), CB_GETCURSEL, 0, 0);
                if (index >= 0 && index <= static_cast<LRESULT>(std::numeric_limits<int>::max()))
                {
                    const size_t choiceIndex = static_cast<size_t>(index);
                    if (choiceIndex < c.field.choices.size())
                    {
                        selected = c.field.choices[choiceIndex].value;
                    }
                }
            }
            else
            {
                for (size_t i = 0; i < c.choiceButtons.size() && i < c.field.choices.size(); ++i)
                {
                    if (SendMessageW(c.choiceButtons[i].get(), BM_GETCHECK, 0, 0) == BST_CHECKED)
                    {
                        selected = c.field.choices[i].value;
                        break;
                    }
                }
            }

            const std::string utf8 = Utf8FromUtf16(selected);
            yyjson_mut_val* val    = yyjson_mut_strncpy(doc, utf8.c_str(), utf8.size());
            if (! val)
            {
                return {};
            }
            if (! yyjson_mut_obj_add(root, key, val))
            {
                return {};
            }
        }
        else if (c.field.type == PrefsPluginConfigFieldType::Selection)
        {
            yyjson_mut_val* arr = yyjson_mut_arr(doc);
            if (! arr)
            {
                return {};
            }
            if (! yyjson_mut_obj_add(root, key, arr))
            {
                return {};
            }

            for (size_t i = 0; i < c.choiceButtons.size() && i < c.field.choices.size(); ++i)
            {
                if (SendMessageW(c.choiceButtons[i].get(), BM_GETCHECK, 0, 0) != BST_CHECKED)
                {
                    continue;
                }

                const std::string utf8 = Utf8FromUtf16(c.field.choices[i].value);
                yyjson_mut_val* val    = yyjson_mut_strncpy(doc, utf8.c_str(), utf8.size());
                if (! val)
                {
                    return {};
                }
                if (! yyjson_mut_arr_add_val(arr, val))
                {
                    return {};
                }
            }
        }
    }

    yyjson_write_err err{};
    size_t len = 0;
    wil::unique_any<char*, decltype(&::free), ::free> json(yyjson_mut_write_opts(doc, YYJSON_WRITE_NOFLAG, nullptr, &len, &err));
    if (! json || len == 0)
    {
        Debug::Error(L"Failed to serialize plugin configuration to JSON: code: {}", err.code);
        return {};
    }

    std::string out(json.get(), len);
    return out;
}

[[nodiscard]] bool ContainsChoiceValue(const std::vector<std::wstring>& values, std::wstring_view needle) noexcept
{
    return std::find(values.begin(), values.end(), needle) != values.end();
}

[[nodiscard]] size_t FindChoiceIndex(const std::vector<PrefsPluginConfigChoice>& choices, std::wstring_view desired) noexcept
{
    for (size_t i = 0; i < choices.size(); ++i)
    {
        if (choices[i].value == desired)
        {
            return i;
        }
    }
    return static_cast<size_t>(std::numeric_limits<size_t>::max());
}

[[nodiscard]] std::wstring GetPluginConfigurationSchemaErrorText(const PrefsPluginListItem& pluginItem) noexcept
{
    if (! PrefsPlugins::IsLoadable(pluginItem))
    {
        return LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_SCHEMA_NOT_LOADABLE);
    }
    return LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_SCHEMA_UNAVAILABLE);
}

[[nodiscard]] PrefsPluginConfigFieldControls* FindFieldForControl(PreferencesDialogState& state, HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return nullptr;
    }

    for (PrefsPluginConfigFieldControls& controls : state.pluginsDetailsConfigFields)
    {
        if (controls.edit.get() == hwnd || controls.combo.get() == hwnd || controls.toggle.get() == hwnd || controls.browseButton.get() == hwnd)
        {
            return &controls;
        }

        for (const auto& button : controls.choiceButtons)
        {
            if (button.get() == hwnd)
            {
                return &controls;
            }
        }
    }

    return nullptr;
}

[[nodiscard]] bool CommitEditor(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host || state.pluginsDetailsConfigPluginId.empty() || state.pluginsDetailsConfigFields.empty())
    {
        return false;
    }

    const std::string configJson = BuildConfigurationJson(state.pluginsDetailsConfigFields);
    if (configJson.empty())
    {
        return false;
    }

    Common::Settings::JsonValue parsedValue;
    const HRESULT parseHr = Common::Settings::ParseJsonValue(configJson, parsedValue);
    if (FAILED(parseHr))
    {
        return false;
    }

    bool clearValue = std::holds_alternative<std::monostate>(parsedValue.value);
    if (! clearValue)
    {
        const auto* obj = std::get_if<Common::Settings::JsonValue::ObjectPtr>(&parsedValue.value);
        clearValue      = obj && *obj && (*obj)->members.empty();
    }

    if (clearValue)
    {
        state.workingSettings.plugins.configurationByPluginId.erase(state.pluginsDetailsConfigPluginId);
    }
    else
    {
        state.workingSettings.plugins.configurationByPluginId[state.pluginsDetailsConfigPluginId] = std::move(parsedValue);
    }

    if (HWND dlg = GetParent(host))
    {
        SetDirty(dlg, state);
    }

    return true;
}
} // namespace
