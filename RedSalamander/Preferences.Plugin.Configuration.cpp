// Preferences.Plugin.Configuration.cpp

#include "Framework.h"

#include "Preferences.Plugin.Configuration.h"

#include <algorithm>
#include <array>
#include <cerrno>
#include <cwchar>
#include <limits>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include <commdlg.h>
#include <shobjidl.h>

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "FileSystemPluginManager.h"
#include "Helpers.h"
#include "UiMetrics.h"
#include "ViewerPluginManager.h"

#include "resource.h"

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::Checkbox;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::Toggle;

#ifdef ENABLE_TESTS
std::mutex g_debugPluginConfigurationBrowseResultMutex;

enum class DebugPluginConfigurationBrowseResultKind : uint8_t
{
    Path,
    Cancel,
};

struct DebugPluginConfigurationBrowseResult
{
    DebugPluginConfigurationBrowseResultKind kind = DebugPluginConfigurationBrowseResultKind::Path;
    std::filesystem::path path;
};

std::optional<DebugPluginConfigurationBrowseResult> g_debugNextPluginConfigurationBrowseResult;
#endif

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

#ifdef ENABLE_TESTS
    {
        std::scoped_lock lock(g_debugPluginConfigurationBrowseResultMutex);
        if (g_debugNextPluginConfigurationBrowseResult.has_value())
        {
            const DebugPluginConfigurationBrowseResult result = *g_debugNextPluginConfigurationBrowseResult;
            g_debugNextPluginConfigurationBrowseResult.reset();
            if (result.kind == DebugPluginConfigurationBrowseResultKind::Cancel)
            {
                return false;
            }

            outPath = result.path;
            return ! outPath.empty();
        }
    }
#endif

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

#ifdef ENABLE_TESTS
[[nodiscard]] bool DebugSetPluginConfigurationNextBrowsePathImpl(const std::wstring_view path) noexcept
{
    std::scoped_lock lock(g_debugPluginConfigurationBrowseResultMutex);
    if (path.empty())
    {
        g_debugNextPluginConfigurationBrowseResult.reset();
        return true;
    }

    g_debugNextPluginConfigurationBrowseResult = DebugPluginConfigurationBrowseResult{
        .kind = DebugPluginConfigurationBrowseResultKind::Path,
        .path = std::filesystem::path(path),
    };
    return true;
}

[[nodiscard]] bool DebugCancelPluginConfigurationNextBrowseImpl() noexcept
{
    std::scoped_lock lock(g_debugPluginConfigurationBrowseResultMutex);
    g_debugNextPluginConfigurationBrowseResult = DebugPluginConfigurationBrowseResult{
        .kind = DebugPluginConfigurationBrowseResultKind::Cancel,
        .path = {},
    };
    return true;
}
#endif

[[maybe_unused]] [[nodiscard]] std::wstring SanitizeIntegerText(std::wstring_view text, const bool allowNegative) noexcept
{
    std::wstring sanitized;
    sanitized.reserve(text.size());

    bool sawDigit    = false;
    bool sawSign     = false;
    const bool allow = allowNegative;

    for (const wchar_t ch : text)
    {
        if (ch >= L'0' && ch <= L'9')
        {
            sanitized.push_back(ch);
            sawDigit = true;
            continue;
        }

        if (allow && ! sawSign && ! sawDigit && ch == L'-')
        {
            sanitized.push_back(ch);
            sawSign = true;
        }
    }

    return sanitized;
}

void SyncDxStaticText(PrefsPluginConfigFieldControls& controls) noexcept
{
    if (controls.dxLabelControl)
    {
        controls.dxLabelControl->SetText(controls.field.label);
    }
    if (controls.dxDescriptionControl)
    {
        controls.dxDescriptionControl->SetText(controls.field.description);
    }
}

void SyncDxInteractiveState(PreferencesDialogState& state, PrefsPluginConfigFieldControls& controls) noexcept
{
    const bool previousRefreshing = state.refreshingPluginsPage;
    state.refreshingPluginsPage   = true;
    const auto restoreRefreshing  = wil::scope_exit([&state, previousRefreshing]() noexcept { state.refreshingPluginsPage = previousRefreshing; });

    if (controls.dxEditControl)
    {
        controls.dxEditControl->SetText(controls.retainedText);
    }

    if (controls.dxToggleControl)
    {
        bool checked = controls.retainedToggleValue;
        if (controls.field.type == PrefsPluginConfigFieldType::Option)
        {
            checked = false;
            for (size_t i = 0; i < controls.field.choices.size(); ++i)
            {
                if (controls.field.choices[i].value == controls.retainedOptionValue)
                {
                    checked = (i == controls.toggleOnChoiceIndex);
                    break;
                }
            }
        }
        controls.dxToggleControl->SetChecked(checked);
    }

    if (controls.dxComboControl)
    {
        std::optional<size_t> selectedIndex;
        if (! controls.retainedOptionValue.empty())
        {
            for (size_t i = 0; i < controls.field.choices.size(); ++i)
            {
                if (controls.field.choices[i].value == controls.retainedOptionValue)
                {
                    selectedIndex = i;
                    break;
                }
            }
        }

        controls.dxComboControl->SetSelectedIndex(selectedIndex);
        if (selectedIndex.has_value() && selectedIndex.value() < controls.field.choices.size())
        {
            const auto& choice = controls.field.choices[selectedIndex.value()];
            controls.dxComboControl->SetText(choice.label.empty() ? choice.value : choice.label);
        }
    }

    for (size_t i = 0; i < controls.dxChoiceControls.size() && i < controls.field.choices.size(); ++i)
    {
        if (controls.dxChoiceControls[i].checkbox)
        {
            const auto it = std::find(controls.retainedSelectionValues.begin(), controls.retainedSelectionValues.end(), controls.field.choices[i].value);
            controls.dxChoiceControls[i].checkbox->SetChecked(it != controls.retainedSelectionValues.end());
        }
    }
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

    return OrdinalString::EqualsNoCase(a, b);
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

        bool uiHidden = false;
        if (TryGetBoolValue(item, "x-ui-hidden", uiHidden))
        {
            field.uiHidden = uiHidden;
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

    if (out.field.type == PrefsPluginConfigFieldType::Text)
    {
        out.retainedText = out.field.defaultText;
    }
    else if (out.field.type == PrefsPluginConfigFieldType::Value)
    {
        out.retainedText = std::to_wstring(out.field.defaultInt);
    }
    else if (out.field.type == PrefsPluginConfigFieldType::Bool)
    {
        out.retainedToggleValue = out.field.defaultBool;
    }
    else if (out.field.type == PrefsPluginConfigFieldType::Option)
    {
        out.retainedOptionValue = out.field.defaultOption;
    }
    else if (out.field.type == PrefsPluginConfigFieldType::Selection)
    {
        out.retainedSelectionValues = out.field.defaultSelection;
    }
}
} // namespace

namespace
{
[[nodiscard]] std::wstring GetPluginConfigurationSchemaErrorText(const PrefsPluginListItem& pluginItem) noexcept;
[[nodiscard]] bool CommitEditor(HWND host, PreferencesDialogState& state) noexcept;
} // namespace

namespace PrefsPluginConfiguration
{
void SetDetailsIdText(PreferencesDialogState& state, std::wstring_view text) noexcept
{
    const std::wstring updated(text);
    if (state.pluginsDetailsIdText != updated)
    {
        state.pluginsDetailsIdText = updated;
    }
}

void SetDetailsConfigErrorText(PreferencesDialogState& state, std::wstring_view text) noexcept
{
    const std::wstring updated(text);
    if (state.pluginsDetailsConfigErrorText != updated)
    {
        state.pluginsDetailsConfigErrorText = updated;
    }
    state.pluginsDetailsMessageKind = updated.empty() ? PrefsPluginDetailsMessageKind::None : PrefsPluginDetailsMessageKind::Error;
}

void SetDetailsConfigEmptyStateText(PreferencesDialogState& state, std::wstring_view text) noexcept
{
    const std::wstring updated(text);
    if (state.pluginsDetailsConfigEmptyStateText != updated)
    {
        state.pluginsDetailsConfigEmptyStateText = updated;
    }
    state.pluginsDetailsMessageKind = updated.empty() ? PrefsPluginDetailsMessageKind::None : PrefsPluginDetailsMessageKind::EmptyState;
}

void Clear(PreferencesDialogState& state) noexcept
{
    if (state.pluginsDetailsConfigDxPanel)
    {
        state.pluginsDetailsConfigDxPanel->ClearChildren();
        state.pluginsDetailsConfigDxPanel->SetVisible(false);
    }

    for (PrefsPluginConfigFieldControls& controls : state.pluginsDetailsConfigFields)
    {
        controls.dxLabelControl        = nullptr;
        controls.dxDescriptionControl  = nullptr;
        controls.dxEditControl         = nullptr;
        controls.dxBrowseButtonControl = nullptr;
        controls.dxComboControl        = nullptr;
        controls.dxToggleControl       = nullptr;
        for (auto& choiceControl : controls.dxChoiceControls)
        {
            choiceControl.checkbox = nullptr;
        }
        controls.dxChoiceControls.clear();
    }
    state.pluginsDetailsConfigFields.clear();
    SetDetailsConfigErrorText(state, L"");
    SetDetailsConfigEmptyStateText(state, L"");
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
        bool valid = true;

        for (const PrefsPluginConfigFieldControls& controls : state.pluginsDetailsConfigFields)
        {
            for (const auto& choiceControl : controls.dxChoiceControls)
            {
                valid = valid && (choiceControl.checkbox != nullptr);
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
        SetDetailsConfigErrorText(state, GetPluginConfigurationSchemaErrorText(pluginItem));
        return false;
    }

    const std::vector<PrefsPluginConfigField> fields = ParsePluginConfigSchema(schemaUtf8);
    if (fields.empty())
    {
        SetDetailsConfigEmptyStateText(state, LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_SCHEMA_NO_FIELDS));
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
    auto freeConfigDoc    = wil::scope_exit([&]
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

    SetDetailsConfigErrorText(state, L"");

    const HWND panel = parent;
    Panel* dxPanel   = nullptr;
    if (state.pageHostDxContentRootControl)
    {
        if (! state.pluginsDetailsConfigDxPanel)
        {
            state.pluginsDetailsConfigDxPanel = state.pageHostDxContentRootControl->AddChild<Panel>();
        }

        dxPanel = state.pluginsDetailsConfigDxPanel;
        dxPanel->SetVisible(true);
        dxPanel->ClearChildren();
    }

    state.pluginsDetailsConfigFields.clear();
    state.pluginsDetailsConfigFields.reserve(fields.size());

    for (const PrefsPluginConfigField& field : fields)
    {
        PrefsPluginConfigFieldControls controls{};
        ApplyFieldDefaultToControls(field, controls, configRoot);

        if (controls.field.uiHidden)
        {
            state.pluginsDetailsConfigFields.push_back(std::move(controls));
            continue;
        }

        if (! controls.field.uiHidden && dxPanel)
        {
            controls.dxLabelControl = dxPanel->AddChild<Label>();
            controls.dxLabelControl->SetFontRole(FontRole::Body);
            controls.dxDescriptionControl = dxPanel->AddChild<Label>();
            controls.dxDescriptionControl->SetFontRole(FontRole::Small);
            controls.dxDescriptionControl->SetMultiline(true);
            SyncDxStaticText(controls);
        }

        if (! controls.field.uiHidden && dxPanel)
        {
            if (controls.field.type == PrefsPluginConfigFieldType::Text || controls.field.type == PrefsPluginConfigFieldType::Value)
            {
                const bool numericOnly = controls.field.type == PrefsPluginConfigFieldType::Value;
                controls.dxEditControl = dxPanel->AddChild<TextField>();
                if (controls.dxEditControl)
                {
                    auto* fieldControl          = controls.dxEditControl;
                    const std::wstring fieldKey = controls.field.key;
                    fieldControl->SetOnTextChanged([parent = panel, fieldKey, fieldControl, numericOnly](std::wstring_view text) noexcept
                    {
                        if (! parent || ! fieldControl || IsWindow(parent) == FALSE)
                        {
                            return;
                        }

                        auto* state = PrefsUi::GetDialogState(parent);
                        if (! state || state->refreshingPluginsPage)
                        {
                            return;
                        }

                        std::wstring normalized = numericOnly ? SanitizeIntegerText(text, false) : std::wstring(text);
                        if (normalized != text)
                        {
                            const bool previousRefreshing = state->refreshingPluginsPage;
                            state->refreshingPluginsPage  = true;
                            fieldControl->SetText(normalized);
                            state->refreshingPluginsPage = previousRefreshing;
                        }

                        auto controls = std::find_if(state->pluginsDetailsConfigFields.begin(),
                                                     state->pluginsDetailsConfigFields.end(),
                                                     [&](const PrefsPluginConfigFieldControls& candidate) noexcept { return candidate.field.key == fieldKey; });
                        if (controls == state->pluginsDetailsConfigFields.end() || controls->retainedText == normalized)
                        {
                            return;
                        }

                        controls->retainedText = std::move(normalized);
                    });
                    fieldControl->SetOnBlur([parent = panel]() noexcept
                    {
                        auto* state = PrefsUi::GetDialogState(parent);
                        if (! state || state->refreshingPluginsPage)
                        {
                            return;
                        }

                        static_cast<void>(CommitEditor(parent, *state));
                    });
                    fieldControl->SetOnSubmitted([parent = panel]() noexcept
                    {
                        auto* state = PrefsUi::GetDialogState(parent);
                        if (! state || state->refreshingPluginsPage)
                        {
                            return;
                        }

                        static_cast<void>(CommitEditor(parent, *state));
                    });
                }
            }

            if (controls.field.type == PrefsPluginConfigFieldType::Text && controls.field.browseFolder)
            {
                const std::wstring label       = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_CONFIG_BROWSE_ELLIPSIS);
                controls.dxBrowseButtonControl = dxPanel->AddChild<Button>();
                if (controls.dxBrowseButtonControl)
                {
                    controls.dxBrowseButtonControl->SetText(label);
                    const std::wstring fieldKey = controls.field.key;
                    controls.dxBrowseButtonControl->SetOnClick([parent = panel, fieldKey]() noexcept
                    {
                        auto* state = PrefsUi::GetDialogState(parent);
                        if (! state || state->refreshingPluginsPage || IsWindow(parent) == FALSE)
                        {
                            return;
                        }

                        auto controls = std::find_if(state->pluginsDetailsConfigFields.begin(),
                                                     state->pluginsDetailsConfigFields.end(),
                                                     [&](const PrefsPluginConfigFieldControls& candidate) noexcept { return candidate.field.key == fieldKey; });
                        if (controls == state->pluginsDetailsConfigFields.end() || ! controls->field.browseFolder)
                        {
                            return;
                        }

                        std::filesystem::path selectedPath;
                        HWND owner = GetParent(parent);
                        if (! TryBrowseFolderPath(owner ? owner : parent, selectedPath))
                        {
                            return;
                        }

                        controls->retainedText        = selectedPath.wstring();
                        const bool previousRefreshing = state->refreshingPluginsPage;
                        state->refreshingPluginsPage  = true;
                        if (controls->dxEditControl)
                        {
                            controls->dxEditControl->SetText(controls->retainedText);
                        }
                        state->refreshingPluginsPage = previousRefreshing;
                        static_cast<void>(CommitEditor(parent, *state));
                    });
                }
            }

            if (controls.field.type == PrefsPluginConfigFieldType::Option && controls.field.choices.size() >= 2u && ! state.theme.systemHighContrast)
            {
                std::wstring uncheckedLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
                std::wstring checkedLabel   = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
                const auto& onChoice        = controls.field.choices[std::min(controls.toggleOnChoiceIndex, controls.field.choices.size() - 1u)];
                const auto& offChoice       = controls.field.choices[std::min(controls.toggleOffChoiceIndex, controls.field.choices.size() - 1u)];
                checkedLabel                = onChoice.label.empty() ? onChoice.value : onChoice.label;
                uncheckedLabel              = offChoice.label.empty() ? offChoice.value : offChoice.label;

                controls.dxToggleControl = dxPanel->AddChild<Toggle>();
                if (controls.dxToggleControl)
                {
                    controls.dxToggleControl->SetStateLabels(std::move(uncheckedLabel), std::move(checkedLabel));
                    const std::wstring fieldKey = controls.field.key;
                    controls.dxToggleControl->SetOnToggled([parent = panel, fieldKey](bool checked) noexcept
                    {
                        auto* state = PrefsUi::GetDialogState(parent);
                        if (! state || state->refreshingPluginsPage)
                        {
                            return;
                        }

                        auto controls = std::find_if(state->pluginsDetailsConfigFields.begin(),
                                                     state->pluginsDetailsConfigFields.end(),
                                                     [&](const PrefsPluginConfigFieldControls& candidate) noexcept { return candidate.field.key == fieldKey; });
                        if (controls == state->pluginsDetailsConfigFields.end() || controls->field.choices.empty())
                        {
                            return;
                        }

                        const size_t choiceIndex = checked ? controls->toggleOnChoiceIndex : controls->toggleOffChoiceIndex;
                        if (choiceIndex < controls->field.choices.size())
                        {
                            controls->retainedOptionValue = controls->field.choices[choiceIndex].value;
                            static_cast<void>(CommitEditor(parent, *state));
                        }
                    });
                }
            }
            else if (controls.field.type == PrefsPluginConfigFieldType::Option)
            {
                std::vector<ComboBox::Item> items;
                items.reserve(controls.field.choices.size());
                for (const auto& choice : controls.field.choices)
                {
                    const std::wstring_view label = choice.label.empty() ? std::wstring_view(choice.value) : std::wstring_view(choice.label);
                    items.push_back({std::wstring(label), choice.value});
                }

                controls.dxComboControl = dxPanel->AddChild<ComboBox>();
                if (controls.dxComboControl)
                {
                    controls.dxComboControl->SetVariant(ComboBoxVariant::Window);
                    controls.dxComboControl->SetItems(std::move(items));
                    const std::wstring fieldKey = controls.field.key;
                    controls.dxComboControl->SetOnSelectionChanged([parent = panel, fieldKey](size_t itemIndex) noexcept
                    {
                        if (! parent || IsWindow(parent) == FALSE)
                        {
                            return;
                        }

                        auto* state = PrefsUi::GetDialogState(parent);
                        if (! state || state->refreshingPluginsPage)
                        {
                            return;
                        }

                        auto controls = std::find_if(state->pluginsDetailsConfigFields.begin(),
                                                     state->pluginsDetailsConfigFields.end(),
                                                     [&](const PrefsPluginConfigFieldControls& candidate) noexcept { return candidate.field.key == fieldKey; });
                        if (controls == state->pluginsDetailsConfigFields.end() || itemIndex >= controls->field.choices.size())
                        {
                            return;
                        }

                        controls->retainedOptionValue = controls->field.choices[itemIndex].value;
                        static_cast<void>(CommitEditor(parent, *state));
                    });
                }
            }

            if (controls.field.type == PrefsPluginConfigFieldType::Bool)
            {
                std::wstring uncheckedLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
                std::wstring checkedLabel   = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);

                controls.dxToggleControl = dxPanel->AddChild<Toggle>();
                if (controls.dxToggleControl)
                {
                    controls.dxToggleControl->SetStateLabels(std::move(uncheckedLabel), std::move(checkedLabel));
                    const std::wstring fieldKey = controls.field.key;
                    controls.dxToggleControl->SetOnToggled([parent = panel, fieldKey](bool checked) noexcept
                    {
                        auto* state = PrefsUi::GetDialogState(parent);
                        if (! state || state->refreshingPluginsPage)
                        {
                            return;
                        }

                        auto controls = std::find_if(state->pluginsDetailsConfigFields.begin(),
                                                     state->pluginsDetailsConfigFields.end(),
                                                     [&](const PrefsPluginConfigFieldControls& candidate) noexcept { return candidate.field.key == fieldKey; });
                        if (controls == state->pluginsDetailsConfigFields.end())
                        {
                            return;
                        }

                        controls->retainedToggleValue = checked;
                        static_cast<void>(CommitEditor(parent, *state));
                    });
                }
            }

            if (controls.field.type == PrefsPluginConfigFieldType::Selection && ! controls.field.choices.empty())
            {
                controls.dxChoiceControls.clear();
                controls.dxChoiceControls.resize(controls.field.choices.size());

                for (size_t i = 0; i < controls.field.choices.size(); ++i)
                {
                    const auto& choice      = controls.field.choices[i];
                    const std::wstring text = choice.label.empty() ? choice.value : choice.label;
                    auto& dxChoice          = controls.dxChoiceControls[i];
                    dxChoice.checkbox       = dxPanel->AddChild<Checkbox>();
                    if (dxChoice.checkbox)
                    {
                        dxChoice.checkbox->SetText(text);
                        const std::wstring fieldKey = controls.field.key;
                        dxChoice.checkbox->SetOnToggled([parent = panel, fieldKey, choiceValue = choice.value](bool checked) noexcept
                        {
                            auto* state = PrefsUi::GetDialogState(parent);
                            if (! state || state->refreshingPluginsPage)
                            {
                                return;
                            }

                            auto controls =
                                std::find_if(state->pluginsDetailsConfigFields.begin(),
                                             state->pluginsDetailsConfigFields.end(),
                                             [&](const PrefsPluginConfigFieldControls& candidate) noexcept { return candidate.field.key == fieldKey; });
                            if (controls == state->pluginsDetailsConfigFields.end())
                            {
                                return;
                            }

                            auto& values  = controls->retainedSelectionValues;
                            const auto it = std::find(values.begin(), values.end(), choiceValue);
                            if (checked)
                            {
                                if (it == values.end())
                                {
                                    values.push_back(choiceValue);
                                }
                            }
                            else if (it != values.end())
                            {
                                values.erase(it);
                            }

                            static_cast<void>(CommitEditor(parent, *state));
                        });
                    }
                }
            }

            SyncDxInteractiveState(state, controls);
        }

        if (controls.dxLabelControl)
        {
            if (controls.dxEditControl)
            {
                controls.dxLabelControl->SetMnemonicTarget(controls.dxEditControl);
            }
            else if (controls.dxComboControl)
            {
                controls.dxLabelControl->SetMnemonicTarget(controls.dxComboControl);
            }
            else if (controls.dxToggleControl)
            {
                controls.dxLabelControl->SetMnemonicTarget(controls.dxToggleControl);
            }
            else if (! controls.dxChoiceControls.empty() && controls.dxChoiceControls.front().checkbox)
            {
                controls.dxLabelControl->SetMnemonicTarget(controls.dxChoiceControls.front().checkbox);
            }
        }

        if (! controls.field.uiHidden)
        {
            SyncDxStaticText(controls);
        }

        state.pluginsDetailsConfigFields.push_back(std::move(controls));
    }

    const bool hasVisibleField = std::any_of(state.pluginsDetailsConfigFields.begin(),
                                             state.pluginsDetailsConfigFields.end(),
                                             [](const PrefsPluginConfigFieldControls& controls) noexcept { return ! controls.field.uiHidden; });
    if (! hasVisibleField)
    {
        SetDetailsConfigEmptyStateText(state, LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_SCHEMA_NO_FIELDS));
    }

    return hasVisibleField;
}

void LayoutCards(HWND host, PreferencesDialogState& state, int x, int& y, int width, const PreferencesTypographyContext& typography) noexcept
{
    if (! host || width <= 0)
    {
        return;
    }

    Debug::Perf::Scope layoutPerf(L"preferences.ui.plugin_configuration_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(static_cast<uint64_t>(typography.dpi));

    const UINT dpi = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);

    const int rowHeight      = std::max(1, UiMetrics::ScaleDip(dpi, 26));
    const int titleHeight    = std::max(1, UiMetrics::ScaleDip(dpi, 18));
    const int optionHeight   = std::max(1, UiMetrics::ScaleDip(dpi, 20));
    const int minToggleWidth = UiMetrics::ScaleDip(dpi, 90);

    const int cardPaddingX = UiMetrics::ScaleDip(dpi, 12);
    const int cardPaddingY = UiMetrics::ScaleDip(dpi, 8);
    const int cardGapY     = UiMetrics::ScaleDip(dpi, 2);
    const int cardGapX     = UiMetrics::ScaleDip(dpi, 12);
    const int cardSpacingY = UiMetrics::ScaleDip(dpi, 8);
    const int innerGapX    = UiMetrics::ScaleDip(dpi, 8);
    const int minInfoWidth = UiMetrics::ScaleDip(dpi, 220);

    const int maxControlWidth = std::max(0, width - 2 * cardPaddingX);

    const auto pushCard = [&](const RECT& card) noexcept { state.pageSettingCards.push_back(card); };

    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
    const auto pxToDip          = [dpi](const int value) noexcept { return (static_cast<float>(value) * 96.0f) / static_cast<float>(std::max<UINT>(1u, dpi)); };
    Panel* const dxPanel        = state.pluginsDetailsConfigDxPanel;
    for (PrefsPluginConfigFieldControls& controls : state.pluginsDetailsConfigFields)
    {
        if (controls.field.uiHidden)
        {
            continue;
        }

        SyncDxStaticText(controls);

        SyncDxInteractiveState(state, controls);

        const bool isSelection = controls.field.type == PrefsPluginConfigFieldType::Selection;

        const std::wstring& descText = controls.field.description;
        const bool hasDesc           = ! descText.empty();

        int controlGroupWidth = 0;
        int editWidth         = 0;
        int browseWidth       = 0;

        if (! isSelection)
        {
            if (controls.dxEditControl)
            {
                const int minEditWidth = UiMetrics::ScaleDip(dpi, 140);
                int desiredWidth       = minEditWidth;

                if (controls.field.type == PrefsPluginConfigFieldType::Text)
                {
                    desiredWidth = UiMetrics::ScaleDip(dpi, controls.field.browseFolder ? 380 : 320);
                }

                if (controls.field.type == PrefsPluginConfigFieldType::Value)
                {
                    desiredWidth = UiMetrics::ScaleDip(dpi, 140);
                }

                browseWidth = controls.dxBrowseButtonControl ? UiMetrics::ScaleDip(dpi, 90) : 0;
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
            else if (controls.dxComboControl)
            {
                int desiredWidth  = UiMetrics::ScaleDip(dpi, 220);
                desiredWidth      = std::max(desiredWidth, UiMetrics::ScaleDip(dpi, 160));
                desiredWidth      = std::min(desiredWidth, std::min(maxControlWidth, UiMetrics::ScaleDip(dpi, 260)));
                controlGroupWidth = desiredWidth;
            }
            else if (controls.dxToggleControl)
            {
                int desiredWidth = std::min(maxControlWidth, UiMetrics::ScaleDip(dpi, 180));
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

                    const int paddingX   = UiMetrics::ScaleDip(dpi, 6);
                    const int gapX       = UiMetrics::ScaleDip(dpi, 8);
                    const int trackWidth = UiMetrics::ScaleDip(dpi, 34);

                    const int onWidth        = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, onStateLabel);
                    const int offWidth       = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, offStateLabel);
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

        if (! isSelection && controlGroupWidth > 0)
        {
            const int maxControlGroupWidth = std::max(1, maxControlWidth - cardGapX - minInfoWidth);
            if (controlGroupWidth > maxControlGroupWidth)
            {
                controlGroupWidth = maxControlGroupWidth;
                if (controls.dxEditControl)
                {
                    const int browseExtra = (browseWidth > 0) ? (innerGapX + browseWidth) : 0;
                    editWidth             = std::max(1, controlGroupWidth - browseExtra);
                    controlGroupWidth     = editWidth + browseExtra;
                }
            }
        }

        const int textWidth = std::max(0, width - 2 * cardPaddingX - ((controlGroupWidth > 0) ? (cardGapX + controlGroupWidth) : 0));

        int descHeight = 0;
        if (hasDesc)
        {
            descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);
        }

        int cardHeight = 0;
        if (isSelection)
        {
            const int optionCount   = static_cast<int>(controls.dxChoiceControls.size());
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

        if (controls.dxLabelControl)
        {
            controls.dxLabelControl->SetVisible(true);
            controls.dxLabelControl->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                           pxToDip(card.top + cardPaddingY),
                                                           pxToDip(card.left + cardPaddingX + textWidth),
                                                           pxToDip(card.top + cardPaddingY + titleHeight)));
        }

        if (! isSelection)
        {
            if (controls.dxEditControl)
            {
                controls.dxEditControl->SetVisible(true);
                controls.dxEditControl->SetBounds(
                    D2D1::RectF(pxToDip(controlX), pxToDip(controlY), pxToDip(controlX + editWidth), pxToDip(controlY + rowHeight)));
            }

            if (controls.dxBrowseButtonControl)
            {
                if (browseWidth > 0)
                {
                    controls.dxBrowseButtonControl->SetVisible(true);
                    controls.dxBrowseButtonControl->SetBounds(D2D1::RectF(pxToDip(controlX + editWidth + innerGapX),
                                                                          pxToDip(controlY),
                                                                          pxToDip(controlX + editWidth + innerGapX + browseWidth),
                                                                          pxToDip(controlY + rowHeight)));
                }
                else
                {
                    controls.dxBrowseButtonControl->SetVisible(false);
                }
            }

            if (controls.dxComboControl)
            {
                controls.dxComboControl->SetVisible(true);
                controls.dxComboControl->SetBounds(
                    D2D1::RectF(pxToDip(controlX), pxToDip(controlY), pxToDip(controlX + controlGroupWidth), pxToDip(controlY + rowHeight)));
            }
            if (controls.dxToggleControl)
            {
                controls.dxToggleControl->SetVisible(true);
                controls.dxToggleControl->SetBounds(
                    D2D1::RectF(pxToDip(controlX), pxToDip(controlY), pxToDip(controlX + controlGroupWidth), pxToDip(controlY + rowHeight)));
            }
        }
        else
        {
            int contentY = card.top + cardPaddingY + titleHeight;
            if (! controls.dxChoiceControls.empty())
            {
                contentY += cardGapY;
            }

            const int optionWidth    = std::max(0, width - 2 * cardPaddingX);
            const size_t optionCount = controls.dxChoiceControls.size();
            for (size_t i = 0; i < optionCount; ++i)
            {
                const int buttonY = contentY + static_cast<int>(i) * optionHeight;
                if (controls.dxChoiceControls[i].checkbox)
                {
                    controls.dxChoiceControls[i].checkbox->SetVisible(true);
                    controls.dxChoiceControls[i].checkbox->SetBounds(D2D1::RectF(
                        pxToDip(card.left + cardPaddingX), pxToDip(buttonY), pxToDip(card.left + cardPaddingX + optionWidth), pxToDip(buttonY + optionHeight)));
                }
            }

            contentY += static_cast<int>(controls.dxChoiceControls.size()) * optionHeight;
            if (hasDesc)
            {
                contentY += cardGapY;
            }
        }

        if (controls.dxDescriptionControl)
        {
            if (hasDesc)
            {
                int descY = 0;
                if (isSelection)
                {
                    descY = card.top + cardPaddingY + titleHeight;
                    if (! controls.dxChoiceControls.empty())
                    {
                        descY += cardGapY + static_cast<int>(controls.dxChoiceControls.size()) * optionHeight;
                    }
                    descY += cardGapY;
                }
                else
                {
                    descY = card.top + cardPaddingY + titleHeight + cardGapY;
                }

                controls.dxDescriptionControl->SetVisible(true);
                controls.dxDescriptionControl->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                                     pxToDip(descY),
                                                                     pxToDip(card.left + cardPaddingX + textWidth),
                                                                     pxToDip(descY + std::max(0, descHeight))));
            }
            else
            {
                controls.dxDescriptionControl->SetVisible(false);
            }
        }

        y += cardHeight + cardSpacingY;
    }

    if (dxPanel)
    {
        RECT client{};
        GetClientRect(host, &client);
        dxPanel->SetVisible(true);
        dxPanel->SetBounds(
            D2D1::RectF(0.0f, 0.0f, pxToDip((std::max)(static_cast<int>(client.right), x + width)), pxToDip((std::max)(static_cast<int>(client.bottom), y))));
    }

    if (state.pageHostDxHost)
    {
        state.pageHostDxHost->Invalidate();
    }
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
            const std::string utf8 = Utf8FromUtf16(c.retainedText);
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
            if (! c.retainedText.empty())
            {
                wchar_t* end           = nullptr;
                errno                  = 0;
                const long long parsed = std::wcstoll(c.retainedText.c_str(), &end, 10);
                if (errno == 0 && end != c.retainedText.c_str())
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
            yyjson_mut_val* val = yyjson_mut_bool(doc, c.retainedToggleValue ? true : false);
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
            const std::string utf8 = Utf8FromUtf16(c.retainedOptionValue);
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

            for (const auto& selectedValue : c.retainedSelectionValues)
            {
                const std::string utf8 = Utf8FromUtf16(selectedValue);
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

[[nodiscard]] std::wstring GetPluginConfigurationSchemaErrorText(const PrefsPluginListItem& pluginItem) noexcept
{
    if (! PrefsPlugins::IsLoadable(pluginItem))
    {
        return LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_SCHEMA_NOT_LOADABLE);
    }
    return LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_SCHEMA_UNAVAILABLE);
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

#ifdef ENABLE_TESTS
bool DebugSetPluginConfigurationNextBrowsePath(const std::wstring_view path) noexcept
{
    return DebugSetPluginConfigurationNextBrowsePathImpl(path);
}

bool DebugCancelPluginConfigurationNextBrowse() noexcept
{
    return DebugCancelPluginConfigurationNextBrowseImpl();
}
#endif
