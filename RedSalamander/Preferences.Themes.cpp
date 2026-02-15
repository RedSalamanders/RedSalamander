// Preferences.Themes.cpp

#include "Framework.h"

#include "Preferences.Themes.h"

#include <algorithm>
#include <array>
#include <cwchar>
#include <cwctype>
#include <filesystem>
#include <format>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include <commctrl.h>
#include <commdlg.h>
#include <uxtheme.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

#include "Helpers.h"
#include "HostServices.h"
#include "ThemedControls.h"
#include "WindowMessages.h"
#include "resource.h"

namespace
{
[[nodiscard]] COLORREF ColorRefFromArgb(uint32_t argb) noexcept
{
    return RGB((argb >> 16) & 0xFFu, (argb >> 8) & 0xFFu, argb & 0xFFu);
}

[[nodiscard]] COLORREF CompositeArgbOnBackground(COLORREF background, uint32_t argb) noexcept
{
    const int alpha = static_cast<int>((argb >> 24) & 0xFFu);
    if (alpha <= 0)
    {
        return background;
    }
    const COLORREF rgb = ColorRefFromArgb(argb);
    if (alpha >= 255)
    {
        return rgb;
    }
    return ThemedControls::BlendColor(background, rgb, alpha, 255);
}

void DrawRoundedColorSwatch(HDC hdc, RECT rc, UINT dpi, const AppTheme& theme, COLORREF background, std::optional<uint32_t> argb, bool enabled) noexcept
{
    if (! hdc || rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return;
    }

    const int width  = std::max(0l, rc.right - rc.left);
    const int height = std::max(0l, rc.bottom - rc.top);
    const int radius = std::max(1, std::min(ThemedControls::ScaleDip(dpi, 4), std::min(width, height) / 2));

    COLORREF border =
        theme.systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT) : ThemedControls::BlendColor(background, theme.menu.text, theme.dark ? 70 : 50, 255);
    COLORREF fill = background;
    if (argb.has_value())
    {
        fill = CompositeArgbOnBackground(background, argb.value());
    }

    if (! enabled && ! theme.highContrast)
    {
        fill   = ThemedControls::BlendColor(background, fill, theme.dark ? 120 : 95, 255);
        border = ThemedControls::BlendColor(background, border, theme.dark ? 120 : 95, 255);
    }

    wil::unique_hbrush brush(CreateSolidBrush(fill));
    wil::unique_hpen pen(CreatePen(PS_SOLID, 1, border));
    if (! brush || ! pen)
    {
        return;
    }

    [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc, brush.get());
    [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc, pen.get());
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, radius, radius);
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
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
    if (text.empty() || text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
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

struct BuiltinThemeOption
{
    std::wstring_view id;
    UINT nameId = 0;
};

constexpr std::array<BuiltinThemeOption, 5> kBuiltinThemeOptions = {{
    {L"builtin/system", IDS_PREFS_THEMES_BASE_SYSTEM},
    {L"builtin/light", IDS_PREFS_THEMES_BASE_LIGHT},
    {L"builtin/dark", IDS_PREFS_THEMES_BASE_DARK},
    {L"builtin/rainbow", IDS_PREFS_THEMES_BASE_RAINBOW},
    {L"builtin/highContrast", IDS_PREFS_THEMES_BASE_HIGH_CONTRAST},
}};

constexpr std::wstring_view kNewThemeComboId = L"__newTheme";

[[nodiscard]] std::wstring GetBuiltinThemeName(std::wstring_view themeId) noexcept
{
    for (const auto& option : kBuiltinThemeOptions)
    {
        if (option.id == themeId)
        {
            return LoadStringResource(nullptr, option.nameId);
        }
    }
    return std::wstring{};
}

void ShowDialogAlert(HWND dlg, HostAlertSeverity severity, const std::wstring& title, const std::wstring& message) noexcept
{
    if (! dlg || message.empty())
    {
        return;
    }

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = severity;
    request.targetWindow = dlg;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(HostShowAlert(request));
}

[[nodiscard]] std::filesystem::path TryGetExecutableDirectory() noexcept
{
    wil::unique_cotaskmem_string modulePath = wil::GetModuleFileNameW();
    if (! modulePath)
    {
        return {};
    }

    std::filesystem::path path(modulePath.get());
    if (! path.has_parent_path())
    {
        return {};
    }

    return path.parent_path();
}

[[nodiscard]] std::filesystem::path TryGetThemesDirectory() noexcept
{
    const std::filesystem::path exeDir = TryGetExecutableDirectory();
    if (exeDir.empty())
    {
        return {};
    }
    return exeDir / L"Themes";
}

[[nodiscard]] bool IsValidThemeColorKey(std::wstring_view key) noexcept
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

[[nodiscard]] bool IsValidUserThemeId(std::wstring_view id) noexcept
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

[[nodiscard]] bool IsBuiltinThemeId(std::wstring_view themeId) noexcept
{
    for (const auto& option : kBuiltinThemeOptions)
    {
        if (option.id == themeId)
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] bool DoesThemeIdExist(const PreferencesDialogState& state, std::wstring_view themeId) noexcept
{
    if (themeId.empty())
    {
        return false;
    }

    if (IsBuiltinThemeId(themeId))
    {
        return true;
    }

    for (const auto& theme : state.workingSettings.theme.themes)
    {
        if (theme.id == themeId)
        {
            return true;
        }
    }

    for (const auto& theme : state.themeFileThemes)
    {
        if (theme.id == themeId)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool DoesThemeIdExistExcluding(const PreferencesDialogState& state, std::wstring_view themeId, std::wstring_view excludedId) noexcept
{
    if (themeId.empty())
    {
        return false;
    }

    if (! excludedId.empty() && themeId == excludedId)
    {
        return false;
    }

    if (IsBuiltinThemeId(themeId))
    {
        return true;
    }

    for (const auto& theme : state.workingSettings.theme.themes)
    {
        if (theme.id == themeId && theme.id != excludedId)
        {
            return true;
        }
    }

    for (const auto& theme : state.themeFileThemes)
    {
        if (theme.id == themeId)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::wstring SlugifyThemeName(std::wstring_view name) noexcept
{
    std::wstring slug;
    slug.reserve(std::min<size_t>(name.size(), 64u));

    bool lastWasSeparator = false;
    for (wchar_t ch : name)
    {
        if (ch >= L'A' && ch <= L'Z')
        {
            slug.push_back(static_cast<wchar_t>(ch - L'A' + L'a'));
            lastWasSeparator = false;
            continue;
        }
        if ((ch >= L'a' && ch <= L'z') || (ch >= L'0' && ch <= L'9'))
        {
            slug.push_back(ch);
            lastWasSeparator = false;
            continue;
        }

        const bool separator = ch == L' ' || ch == L'\t' || ch == L'\r' || ch == L'\n' || ch == L'-' || ch == L'_' || ch == L'.';
        if (! separator)
        {
            continue;
        }

        if (! slug.empty() && ! lastWasSeparator)
        {
            slug.push_back(L'-');
            lastWasSeparator = true;
        }
    }

    while (! slug.empty() && slug.front() == L'-')
    {
        slug.erase(slug.begin());
    }
    while (! slug.empty() && slug.back() == L'-')
    {
        slug.pop_back();
    }

    if (slug.empty())
    {
        return L"theme";
    }

    if (slug.size() > 64u)
    {
        slug.resize(64u);
    }

    const wchar_t first = slug.front();
    const bool firstOk  = (first >= L'a' && first <= L'z') || (first >= L'0' && first <= L'9');
    if (! firstOk)
    {
        slug.insert(slug.begin(), L't');
    }

    if (slug.size() > 64u)
    {
        slug.resize(64u);
    }

    return slug;
}

[[nodiscard]] std::wstring MakeUniqueUserThemeId(PreferencesDialogState& state, std::wstring_view name) noexcept
{
    std::wstring base = SlugifyThemeName(name);
    if (base.empty())
    {
        base = L"theme";
    }

    const auto makeCandidate = [&](std::wstring_view suffix) noexcept -> std::wstring { return std::format(L"user/{}", suffix); };

    std::wstring candidate = makeCandidate(base);
    if (! DoesThemeIdExist(state, candidate))
    {
        return candidate;
    }

    for (int attempt = 2; attempt < 1000; ++attempt)
    {
        std::wstring attemptText;
        attemptText = std::format(L"-{}", attempt);

        std::wstring trimmed      = base;
        const size_t maxSuffixLen = 64u;
        if (attemptText.size() < maxSuffixLen && trimmed.size() > (maxSuffixLen - attemptText.size()))
        {
            trimmed.resize(maxSuffixLen - attemptText.size());
        }

        std::wstring suffix;
        suffix = trimmed + attemptText;

        candidate = makeCandidate(suffix);
        if (! DoesThemeIdExist(state, candidate))
        {
            return candidate;
        }
    }

    return L"user/theme";
}

[[nodiscard]] std::wstring MakeUniqueUserThemeIdForRename(const PreferencesDialogState& state, std::wstring_view name, std::wstring_view existingId) noexcept
{
    std::wstring base = SlugifyThemeName(name);
    if (base.empty())
    {
        base = L"theme";
    }

    const auto makeCandidate = [&](std::wstring_view suffix) noexcept -> std::wstring { return std::format(L"user/{}", suffix); };

    std::wstring candidate = makeCandidate(base);
    if (candidate == existingId)
    {
        return candidate;
    }
    if (! DoesThemeIdExistExcluding(state, candidate, existingId))
    {
        return candidate;
    }

    for (int attempt = 2; attempt < 1000; ++attempt)
    {
        std::wstring attemptText;
        attemptText = std::format(L"-{}", attempt);

        std::wstring trimmed      = base;
        const size_t maxSuffixLen = 64u;
        if (attemptText.size() < maxSuffixLen && trimmed.size() > (maxSuffixLen - attemptText.size()))
        {
            trimmed.resize(maxSuffixLen - attemptText.size());
        }

        std::wstring suffix;
        suffix = trimmed + attemptText;

        candidate = makeCandidate(suffix);
        if (! DoesThemeIdExistExcluding(state, candidate, existingId))
        {
            return candidate;
        }
    }

    return L"user/theme";
}

[[nodiscard]] std::wstring MakeSuggestedThemeFileName(std::wstring_view themeId, std::wstring_view themeName) noexcept
{
    std::wstring base;
    const std::wstring defaultBase = LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_THEME);
    if (! themeName.empty())
    {
        base.assign(themeName);
    }
    else if (themeId.rfind(L"user/", 0) == 0)
    {
        base.assign(themeId.substr(5));
    }
    else
    {
        base.assign(defaultBase.empty() ? std::wstring(themeId) : defaultBase);
    }

    for (auto& ch : base)
    {
        if (ch == L'\\' || ch == L'/' || ch == L':' || ch == L'*' || ch == L'?' || ch == L'\"' || ch == L'<' || ch == L'>' || ch == L'|')
        {
            ch = L'_';
        }
    }

    if (base.empty())
    {
        base.assign(defaultBase.empty() ? std::wstring(themeId) : defaultBase);
    }
    base.append(L".theme.json5");
    return base;
}

[[nodiscard]] bool TryBrowseThemeFile(HWND owner, bool saving, std::wstring_view suggestedFileName, std::filesystem::path& outPath) noexcept
{
    outPath.clear();

    std::array<wchar_t, 1024> buffer{};
    buffer[0] = L'\0';
    if (saving && ! suggestedFileName.empty())
    {
        const size_t copyLen = std::min(suggestedFileName.size(), buffer.size() - 1u);
        std::wmemcpy(buffer.data(), suggestedFileName.data(), copyLen);
        buffer[copyLen] = L'\0';
    }

    const std::wstring filter = LoadStringResource(nullptr, IDS_PREFS_THEMES_FILE_FILTER);

    std::wstring initialDir;
    const std::filesystem::path themesDir = TryGetThemesDirectory();
    if (! themesDir.empty())
    {
        initialDir = themesDir.wstring();
    }

    OPENFILENAMEW ofn{};
    ofn.lStructSize     = sizeof(ofn);
    ofn.hwndOwner       = owner;
    ofn.lpstrFilter     = filter.c_str();
    ofn.lpstrFile       = buffer.data();
    ofn.nMaxFile        = static_cast<DWORD>(buffer.size());
    ofn.lpstrDefExt     = L"json5";
    ofn.lpstrInitialDir = initialDir.empty() ? nullptr : initialDir.c_str();
    ofn.Flags =
        static_cast<DWORD>(OFN_NOCHANGEDIR | OFN_HIDEREADONLY | (saving ? (OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST) : (OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST)));

    const BOOL ok = saving ? GetSaveFileNameW(&ofn) : GetOpenFileNameW(&ofn);
    if (! ok)
    {
        return false;
    }

    outPath = std::filesystem::path(buffer.data());
    return ! outPath.empty();
}

[[nodiscard]] bool ParseThemeDefinitionJson(std::string_view jsonText, Common::Settings::ThemeDefinition& outTheme, std::wstring& outError) noexcept
{
    outError.clear();
    outTheme = {};

    if (jsonText.empty())
    {
        outError = LoadStringResource(nullptr, IDS_PREFS_THEMES_IMPORT_FILE_EMPTY);
        return false;
    }

    std::string buffer(jsonText);
    yyjson_read_err err{};
    wil::unique_any<yyjson_doc*, decltype(&yyjson_doc_free), yyjson_doc_free> doc(
        yyjson_read_opts(buffer.data(), buffer.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err));
    if (! doc)
    {
        const std::wstring msg = (err.msg && err.msg[0] != '\0') ? Utf16FromUtf8(err.msg) : std::wstring{};
        outError               = msg.empty() ? LoadStringResource(nullptr, IDS_PREFS_THEMES_IMPORT_PARSE_FAILED) : msg;
        return false;
    }

    yyjson_val* root = yyjson_doc_get_root(doc.get());
    if (! root || ! yyjson_is_obj(root))
    {
        outError = LoadStringResource(nullptr, IDS_PREFS_THEMES_IMPORT_ROOT_NOT_OBJECT);
        return false;
    }

    const auto requireString = [&](const char* key, std::wstring& dest) -> bool
    {
        yyjson_val* val = yyjson_obj_get(root, key);
        if (! val || ! yyjson_is_str(val))
        {
            outError = FormatStringResource(nullptr, IDS_PREFS_THEMES_IMPORT_FIELD_MISSING_OR_NOT_STRING_FMT, Utf16FromUtf8(std::string_view(key)));
            return false;
        }

        const char* text = yyjson_get_str(val);
        dest             = (text && text[0] != '\0') ? Utf16FromUtf8(text) : std::wstring{};
        if (dest.empty())
        {
            outError = FormatStringResource(nullptr, IDS_PREFS_THEMES_IMPORT_FIELD_EMPTY_FMT, Utf16FromUtf8(std::string_view(key)));
            return false;
        }

        return true;
    };

    if (! requireString("id", outTheme.id))
    {
        return false;
    }
    if (! IsValidUserThemeId(outTheme.id))
    {
        outError = LoadStringResource(nullptr, IDS_PREFS_THEMES_IMPORT_INVALID_ID);
        return false;
    }

    if (! requireString("name", outTheme.name))
    {
        return false;
    }
    if (! requireString("baseThemeId", outTheme.baseThemeId))
    {
        return false;
    }
    if (! IsBuiltinThemeId(outTheme.baseThemeId))
    {
        outError = LoadStringResource(nullptr, IDS_PREFS_THEMES_IMPORT_BASE_NOT_BUILTIN);
        return false;
    }

    yyjson_val* colors = yyjson_obj_get(root, "colors");
    if (! colors || ! yyjson_is_obj(colors))
    {
        outError = LoadStringResource(nullptr, IDS_PREFS_THEMES_IMPORT_COLORS_MISSING_OR_NOT_OBJECT);
        return false;
    }

    yyjson_obj_iter iter = yyjson_obj_iter_with(colors);
    yyjson_val* keyVal   = nullptr;
    while ((keyVal = yyjson_obj_iter_next(&iter)) != nullptr)
    {
        const char* keyText = yyjson_get_str(keyVal);
        if (! keyText || keyText[0] == '\0')
        {
            continue;
        }

        const std::wstring keyWide = Utf16FromUtf8(keyText);
        if (! IsValidThemeColorKey(keyWide))
        {
            continue;
        }

        yyjson_val* valueVal = yyjson_obj_iter_get_val(keyVal);
        if (! valueVal || ! yyjson_is_str(valueVal))
        {
            outError = LoadStringResource(nullptr, IDS_PREFS_THEMES_IMPORT_COLOR_VALUES_MUST_BE_STRINGS);
            return false;
        }

        const char* valueText        = yyjson_get_str(valueVal);
        const std::wstring valueWide = (valueText && valueText[0] != '\0') ? Utf16FromUtf8(valueText) : std::wstring{};
        uint32_t argb                = 0;
        if (valueWide.empty() || ! Common::Settings::TryParseColor(valueWide, argb))
        {
            outError = LoadStringResource(nullptr, IDS_PREFS_THEMES_IMPORT_INVALID_COLOR_VALUE);
            return false;
        }

        outTheme.colors[keyWide] = argb;
    }

    return true;
}

[[nodiscard]] bool BuildThemeDefinitionExportJson(const Common::Settings::ThemeDefinition& theme, std::string& outJson) noexcept
{
    outJson.clear();

    wil::unique_any<yyjson_mut_doc*, decltype(&yyjson_mut_doc_free), yyjson_mut_doc_free> doc(yyjson_mut_doc_new(nullptr));
    if (! doc)
    {
        return false;
    }

    const std::string idUtf8   = Utf8FromUtf16(theme.id);
    const std::string nameUtf8 = Utf8FromUtf16(theme.name);
    const std::string baseUtf8 = Utf8FromUtf16(theme.baseThemeId);
    if (idUtf8.empty() || nameUtf8.empty() || baseUtf8.empty())
    {
        return false;
    }

    yyjson_mut_val* root = yyjson_mut_obj(doc.get());
    if (! root)
    {
        return false;
    }
    yyjson_mut_doc_set_root(doc.get(), root);

    yyjson_mut_val* idVal = yyjson_mut_strncpy(doc.get(), idUtf8.data(), idUtf8.size());
    if (! idVal || ! yyjson_mut_obj_add_val(doc.get(), root, "id", idVal))
    {
        return false;
    }

    yyjson_mut_val* nameVal = yyjson_mut_strncpy(doc.get(), nameUtf8.data(), nameUtf8.size());
    if (! nameVal || ! yyjson_mut_obj_add_val(doc.get(), root, "name", nameVal))
    {
        return false;
    }

    yyjson_mut_val* baseVal = yyjson_mut_strncpy(doc.get(), baseUtf8.data(), baseUtf8.size());
    if (! baseVal || ! yyjson_mut_obj_add_val(doc.get(), root, "baseThemeId", baseVal))
    {
        return false;
    }

    yyjson_mut_val* colors = yyjson_mut_obj(doc.get());
    if (! colors || ! yyjson_mut_obj_add_val(doc.get(), root, "colors", colors))
    {
        return false;
    }

    std::vector<std::wstring_view> keys;
    keys.reserve(theme.colors.size());
    for (const auto& [key, _] : theme.colors)
    {
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
            continue;
        }

        const std::wstring colorText = Common::Settings::FormatColor(it->second);
        const std::string colorUtf8  = Utf8FromUtf16(colorText);
        if (colorUtf8.empty())
        {
            continue;
        }

        yyjson_mut_val* keyVal   = yyjson_mut_strncpy(doc.get(), keyUtf8.data(), keyUtf8.size());
        yyjson_mut_val* valueVal = yyjson_mut_strncpy(doc.get(), colorUtf8.data(), colorUtf8.size());
        if (! keyVal || ! valueVal)
        {
            return false;
        }

        if (! yyjson_mut_obj_add(colors, keyVal, valueVal))
        {
            return false;
        }
    }

    size_t len = 0;
    yyjson_write_err err{};
    wil::unique_any<char*, decltype(&::free), ::free> jsonText(yyjson_mut_write_opts(doc.get(), YYJSON_WRITE_PRETTY, nullptr, &len, &err));
    if (! jsonText || len == 0)
    {
        return false;
    }

    outJson.assign(jsonText.get(), len);
    return ! outJson.empty();
}

void EnsureThemesBaseComboItems(PreferencesDialogState& state) noexcept
{
    if (! state.themesBaseCombo)
    {
        return;
    }

    const LRESULT count = SendMessageW(state.themesBaseCombo.get(), CB_GETCOUNT, 0, 0);
    if (count != CB_ERR && count > 0)
    {
        return;
    }

    SendMessageW(state.themesBaseCombo.get(), CB_RESETCONTENT, 0, 0);
    std::wstring noneText = LoadStringResource(nullptr, IDS_PREFS_THEMES_BASE_NONE);
    if (noneText.empty())
    {
        noneText = LoadStringResource(nullptr, IDS_PREFS_PANES_SORT_NONE);
    }

    const LRESULT noneIdx = SendMessageW(state.themesBaseCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(noneText.empty() ? L"" : noneText.c_str()));
    if (noneIdx != CB_ERR && noneIdx != CB_ERRSPACE)
    {
        SendMessageW(state.themesBaseCombo.get(), CB_SETITEMDATA, static_cast<WPARAM>(noneIdx), static_cast<LPARAM>(-1));
    }
    for (size_t i = 0; i < kBuiltinThemeOptions.size(); ++i)
    {
        const auto& option      = kBuiltinThemeOptions[i];
        const std::wstring name = LoadStringResource(nullptr, option.nameId);
        const LRESULT idx =
            SendMessageW(state.themesBaseCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(name.empty() ? option.id.data() : name.c_str()));
        if (idx != CB_ERR && idx != CB_ERRSPACE)
        {
            SendMessageW(state.themesBaseCombo.get(), CB_SETITEMDATA, static_cast<WPARAM>(idx), static_cast<LPARAM>(i));
        }
    }

    SendMessageW(state.themesBaseCombo.get(), CB_SETCURSEL, 0, 0);
    PrefsUi::InvalidateComboBox(state.themesBaseCombo.get());
}

void EnsureThemeFileThemesLoaded(PreferencesDialogState& state) noexcept
{
    if (! state.themeFileThemes.empty())
    {
        return;
    }

    const std::filesystem::path themesDir = TryGetThemesDirectory();
    if (themesDir.empty())
    {
        return;
    }

    std::vector<Common::Settings::ThemeDefinition> defs;
    const HRESULT hr = Common::Settings::LoadThemeDefinitionsFromDirectory(themesDir, defs);
    if (SUCCEEDED(hr))
    {
        state.themeFileThemes = std::move(defs);
    }
}

void PopulateThemesThemeCombo(PreferencesDialogState& state) noexcept
{
    if (! state.themesThemeCombo)
    {
        return;
    }

    EnsureThemeFileThemesLoaded(state);

    SendMessageW(state.themesThemeCombo.get(), CB_RESETCONTENT, 0, 0);
    state.themeComboItems.clear();

    auto addTheme = [&](std::wstring_view id, std::wstring_view name, ThemeSchemaSource source) noexcept
    {
        ThemeComboItem item;
        item.id          = std::wstring(id);
        item.displayName = name.empty() ? std::wstring(id) : std::wstring(name);
        item.source      = source;

        const LRESULT comboIndex = SendMessageW(state.themesThemeCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item.displayName.c_str()));
        if (comboIndex == CB_ERR || comboIndex == CB_ERRSPACE)
        {
            return;
        }

        state.themeComboItems.push_back(std::move(item));
        SendMessageW(state.themesThemeCombo.get(), CB_SETITEMDATA, static_cast<WPARAM>(comboIndex), static_cast<LPARAM>(state.themeComboItems.size() - 1u));
    };

    for (const auto& builtin : kBuiltinThemeOptions)
    {
        addTheme(builtin.id, LoadStringResource(nullptr, builtin.nameId), ThemeSchemaSource::Builtin);
    }

    auto hasSettingsThemeId = [&](std::wstring_view id) noexcept
    {
        for (const auto& theme : state.workingSettings.theme.themes)
        {
            if (theme.id == id)
            {
                return true;
            }
        }
        return false;
    };

    for (const auto& theme : state.themeFileThemes)
    {
        if (! hasSettingsThemeId(theme.id))
        {
            addTheme(theme.id, theme.name.empty() ? std::wstring_view(theme.id) : std::wstring_view(theme.name), ThemeSchemaSource::File);
        }
    }

    for (const auto& theme : state.workingSettings.theme.themes)
    {
        addTheme(theme.id, theme.name.empty() ? std::wstring_view(theme.id) : std::wstring_view(theme.name), ThemeSchemaSource::Settings);
    }

    addTheme(kNewThemeComboId, LoadStringResource(nullptr, IDS_PREFS_THEMES_NEW_THEME_ENTRY), ThemeSchemaSource::New);

    const std::wstring_view desiredId = state.workingSettings.theme.currentThemeId;
    int desiredIndex                  = 0;
    const LRESULT comboCount          = SendMessageW(state.themesThemeCombo.get(), CB_GETCOUNT, 0, 0);
    for (int i = 0; i < comboCount; ++i)
    {
        const LRESULT data = SendMessageW(state.themesThemeCombo.get(), CB_GETITEMDATA, static_cast<WPARAM>(i), 0);
        if (data == CB_ERR)
        {
            continue;
        }

        const size_t itemIndex = static_cast<size_t>(data);
        if (itemIndex < state.themeComboItems.size() && std::wstring_view(state.themeComboItems[itemIndex].id) == desiredId)
        {
            desiredIndex = i;
            break;
        }
    }

    SendMessageW(state.themesThemeCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(desiredIndex), 0);
    PrefsUi::InvalidateComboBox(state.themesThemeCombo.get());
}

[[nodiscard]] const ThemeComboItem* TryGetSelectedThemeComboItem(const PreferencesDialogState& state) noexcept
{
    if (! state.themesThemeCombo)
    {
        return nullptr;
    }

    const LRESULT sel = SendMessageW(state.themesThemeCombo.get(), CB_GETCURSEL, 0, 0);
    if (sel == CB_ERR)
    {
        return nullptr;
    }

    const LRESULT data = SendMessageW(state.themesThemeCombo.get(), CB_GETITEMDATA, static_cast<WPARAM>(sel), 0);
    if (data == CB_ERR)
    {
        return nullptr;
    }

    const size_t index = static_cast<size_t>(data);
    if (index >= state.themeComboItems.size())
    {
        return nullptr;
    }

    return &state.themeComboItems[index];
}

[[nodiscard]] std::optional<std::wstring_view> TryGetSelectedThemeId(const PreferencesDialogState& state) noexcept
{
    if (const auto* item = TryGetSelectedThemeComboItem(state))
    {
        return std::wstring_view(item->id);
    }
    return std::nullopt;
}

[[nodiscard]] Common::Settings::ThemeDefinition* FindWorkingThemeDefinition(PreferencesDialogState& state, std::wstring_view id) noexcept
{
    for (auto& theme : state.workingSettings.theme.themes)
    {
        if (theme.id == id)
        {
            return &theme;
        }
    }
    return nullptr;
}

[[nodiscard]] const Common::Settings::ThemeDefinition* FindThemeDefinitionById(const std::vector<Common::Settings::ThemeDefinition>& themes,
                                                                               std::wstring_view id) noexcept
{
    for (const auto& theme : themes)
    {
        if (theme.id == id)
        {
            return &theme;
        }
    }
    return nullptr;
}

[[nodiscard]] const Common::Settings::ThemeDefinition*
FindThemeDefinitionForDisplay(const PreferencesDialogState& state, std::wstring_view id, bool& outEditable) noexcept
{
    outEditable = false;
    if (const auto* def = FindThemeDefinitionById(state.workingSettings.theme.themes, id))
    {
        outEditable = true;
        return def;
    }
    if (const auto* def = FindThemeDefinitionById(state.themeFileThemes, id))
    {
        outEditable = false;
        return def;
    }
    return nullptr;
}

[[nodiscard]] ThemeMode ThemeModeFromThemeId(std::wstring_view id) noexcept;
[[nodiscard]] std::optional<D2D1::ColorF> FindAccentOverride(const std::unordered_map<std::wstring, uint32_t>& colors) noexcept;
void ApplyAppThemeOverrides(AppTheme& theme, const std::unordered_map<std::wstring, uint32_t>& colors) noexcept;

struct MonitorTextViewTheme
{
    D2D1::ColorF bg              = D2D1::ColorF(D2D1::ColorF::White);
    D2D1::ColorF fg              = D2D1::ColorF(D2D1::ColorF::Black);
    D2D1::ColorF caret           = D2D1::ColorF(D2D1::ColorF::Black);
    D2D1::ColorF selection       = D2D1::ColorF(0.20f, 0.55f, 0.95f, 0.35f);
    D2D1::ColorF searchHighlight = D2D1::ColorF(1.00f, 0.85f, 0.05f, 0.35f);
    D2D1::ColorF gutterBg        = D2D1::ColorF(D2D1::ColorF::Gainsboro);
    D2D1::ColorF gutterFg        = D2D1::ColorF(D2D1::ColorF::DimGray);
    D2D1::ColorF metaText        = D2D1::ColorF(D2D1::ColorF::DimGray);
    D2D1::ColorF metaError       = D2D1::ColorF(D2D1::ColorF::Red);
    D2D1::ColorF metaWarning     = D2D1::ColorF(D2D1::ColorF::Orange);
    D2D1::ColorF metaInfo        = D2D1::ColorF(D2D1::ColorF::DodgerBlue);
    D2D1::ColorF metaDebug       = D2D1::ColorF(D2D1::ColorF::MediumPurple);
};
[[nodiscard]] MonitorTextViewTheme ResolveMonitorThemeForDisplay(std::wstring_view baseThemeId,
                                                                 const std::unordered_map<std::wstring, uint32_t>* overrides) noexcept;
[[nodiscard]] std::optional<uint32_t> TryGetEffectiveThemeColorArgb(const AppTheme& appTheme,
                                                                    const MonitorTextViewTheme& monitorTheme,
                                                                    const std::unordered_map<std::wstring, uint32_t>* overrides,
                                                                    std::wstring_view key) noexcept;

void EnsureThemesColorsListColumns(HWND list, UINT dpi) noexcept
{
    if (! list)
    {
        return;
    }

    const HWND header  = ListView_GetHeader(list);
    const int existing = header ? Header_GetItemCount(header) : 0;
    if (existing > 0)
    {
        return;
    }

    const std::wstring keyText   = LoadStringResource(nullptr, IDS_PREFS_THEMES_COL_KEY);
    const std::wstring valueText = LoadStringResource(nullptr, IDS_PREFS_THEMES_COL_VALUE);

    LVCOLUMNW col{};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT;
    col.fmt  = LVCFMT_LEFT;

    col.pszText = const_cast<wchar_t*>(keyText.empty() ? L"" : keyText.c_str());
    col.cx      = std::max(0, ThemedControls::ScaleDip(dpi, 260));
    ListView_InsertColumn(list, 0, &col);

    col.pszText = const_cast<wchar_t*>(valueText.empty() ? L"" : valueText.c_str());
    col.cx      = std::max(0, ThemedControls::ScaleDip(dpi, 140));
    ListView_InsertColumn(list, 1, &col);

    col.pszText = const_cast<wchar_t*>(L"");
    col.cx      = std::max(0, ThemedControls::ScaleDip(dpi, 44));
    ListView_InsertColumn(list, 2, &col);
}

void RefreshThemesColorsList(HWND host, PreferencesDialogState& state, std::wstring_view themeId, const Common::Settings::ThemeDefinition* def) noexcept
{
    if (! host || ! state.themesColorsList)
    {
        return;
    }

    const std::wstring_view baseThemeId = (def && ! def->baseThemeId.empty()) ? std::wstring_view(def->baseThemeId) : themeId;
    const auto* overrides               = def ? &def->colors : nullptr;

    const ThemeMode baseMode = ThemeModeFromThemeId(baseThemeId);
    std::optional<D2D1::ColorF> accentOverride;
    if (overrides)
    {
        accentOverride = FindAccentOverride(*overrides);
    }

    AppTheme appTheme = ResolveAppTheme(baseMode, L"RedSalamander", accentOverride);
    if (overrides)
    {
        ApplyAppThemeOverrides(appTheme, *overrides);
    }
    const MonitorTextViewTheme monitorTheme = ResolveMonitorThemeForDisplay(baseThemeId, overrides);

    const UINT dpi = GetDpiForWindow(host);
    EnsureThemesColorsListColumns(state.themesColorsList.get(), dpi);

    std::wstring filterText;
    std::wstring_view filter;
    if (state.themesSearchEdit)
    {
        filterText = PrefsUi::GetWindowTextString(state.themesSearchEdit.get());
        filter     = PrefsUi::TrimWhitespace(filterText);
    }

    std::wstring selectedKey;
    if (const int selected = ListView_GetNextItem(state.themesColorsList.get(), -1, LVNI_SELECTED); selected >= 0)
    {
        wchar_t buffer[128]{};
        ListView_GetItemText(state.themesColorsList.get(), selected, 0, buffer, static_cast<int>(std::size(buffer)));
        selectedKey = buffer;
    }

    ListView_DeleteAllItems(state.themesColorsList.get());

    static constexpr std::array<std::wstring_view, 57> kKnownKeys = {{
        L"app.accent",
        L"window.background",

        L"menu.background",
        L"menu.text",
        L"menu.disabledText",
        L"menu.selectionBg",
        L"menu.selectionText",
        L"menu.separator",
        L"menu.border",

        L"navigation.background",
        L"navigation.backgroundHover",
        L"navigation.backgroundPressed",
        L"navigation.text",
        L"navigation.separator",
        L"navigation.accent",
        L"navigation.progressOk",
        L"navigation.progressWarn",
        L"navigation.progressBackground",

        L"folderView.background",
        L"folderView.itemBackgroundNormal",
        L"folderView.itemBackgroundHovered",
        L"folderView.itemBackgroundSelected",
        L"folderView.itemBackgroundSelectedInactive",
        L"folderView.itemBackgroundFocused",
        L"folderView.textNormal",
        L"folderView.textSelected",
        L"folderView.textSelectedInactive",
        L"folderView.textDisabled",
        L"folderView.focusBorder",
        L"folderView.gridLines",
        L"folderView.errorBackground",
        L"folderView.errorText",
        L"folderView.warningBackground",
        L"folderView.warningText",
        L"folderView.infoBackground",
        L"folderView.infoText",

        L"monitor.textView.bg",
        L"monitor.textView.fg",
        L"monitor.textView.caret",
        L"monitor.textView.selection",
        L"monitor.textView.searchHighlight",
        L"monitor.textView.gutterBg",
        L"monitor.textView.gutterFg",
        L"monitor.textView.metaText",
        L"monitor.textView.metaError",
        L"monitor.textView.metaWarning",
        L"monitor.textView.metaInfo",
        L"monitor.textView.metaDebug",

        L"fileOps.progressBackground",
        L"fileOps.progressTotal",
        L"fileOps.progressItem",
        L"fileOps.graphBackground",
        L"fileOps.graphGrid",
        L"fileOps.graphLimit",
        L"fileOps.graphLine",
        L"fileOps.scrollbarTrack",
        L"fileOps.scrollbarThumb",
    }};

    std::vector<std::wstring> extraKeys;
    if (overrides)
    {
        extraKeys.reserve(overrides->size());
        for (const auto& [key, _] : *overrides)
        {
            bool known = false;
            for (const auto knownKey : kKnownKeys)
            {
                if (knownKey == key)
                {
                    known = true;
                    break;
                }
            }
            if (! known)
            {
                extraKeys.push_back(key);
            }
        }
        std::sort(extraKeys.begin(), extraKeys.end());
    }

    std::vector<std::wstring> allKeys;
    allKeys.reserve(kKnownKeys.size() + extraKeys.size());
    for (const auto key : kKnownKeys)
    {
        allKeys.emplace_back(key);
    }
    for (auto& key : extraKeys)
    {
        allKeys.push_back(std::move(key));
    }

    for (const auto& key : allKeys)
    {
        if (! filter.empty() && ! PrefsUi::ContainsCaseInsensitive(key, filter))
        {
            continue;
        }

        const auto valueOpt = TryGetEffectiveThemeColorArgb(appTheme, monitorTheme, overrides, key);
        if (! valueOpt.has_value())
        {
            continue;
        }

        const std::wstring valueText = Common::Settings::FormatColor(valueOpt.value());
        std::wstring storedKey;
        storedKey.assign(key);

        LVITEMW item{};
        const bool overridden = overrides && overrides->contains(storedKey);

        item.mask       = LVIF_TEXT | LVIF_PARAM;
        item.iItem      = ListView_GetItemCount(state.themesColorsList.get());
        item.iSubItem   = 0;
        item.pszText    = const_cast<wchar_t*>(storedKey.c_str());
        item.lParam     = static_cast<LPARAM>(overridden ? 1 : 0);
        const int index = ListView_InsertItem(state.themesColorsList.get(), &item);
        if (index < 0)
        {
            continue;
        }

        ListView_SetItemText(state.themesColorsList.get(), index, 1, const_cast<wchar_t*>(valueText.c_str()));

        if (! selectedKey.empty() && selectedKey == storedKey)
        {
            ListView_SetItemState(state.themesColorsList.get(), index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        }
    }

    if (ListView_GetNextItem(state.themesColorsList.get(), -1, LVNI_SELECTED) < 0 && ListView_GetItemCount(state.themesColorsList.get()) > 0)
    {
        ListView_SetItemState(state.themesColorsList.get(), 0, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    }
}

void UpdateThemesEnabled(PreferencesDialogState& state, bool editable) noexcept
{
    const BOOL enable = editable ? TRUE : FALSE;

    if (state.themesNameEdit)
    {
        EnableWindow(state.themesNameEdit.get(), enable);
    }
    if (state.themesNameLabel)
    {
        EnableWindow(state.themesNameLabel.get(), enable);
    }
    if (state.themesBaseCombo)
    {
        EnableWindow(state.themesBaseCombo.get(), enable);
    }
    if (state.themesBaseLabel)
    {
        EnableWindow(state.themesBaseLabel.get(), enable);
    }
    if (state.themesColorsList)
    {
        EnableWindow(state.themesColorsList.get(), editable ? TRUE : TRUE);
    }
    if (state.themesKeyEdit)
    {
        EnableWindow(state.themesKeyEdit.get(), enable);
    }
    if (state.themesKeyLabel)
    {
        EnableWindow(state.themesKeyLabel.get(), enable);
    }
    if (state.themesColorEdit)
    {
        EnableWindow(state.themesColorEdit.get(), enable);
    }
    if (state.themesColorLabel)
    {
        EnableWindow(state.themesColorLabel.get(), enable);
    }
    if (state.themesPickColor)
    {
        EnableWindow(state.themesPickColor.get(), enable);
    }
    if (state.themesSetOverride)
    {
        EnableWindow(state.themesSetOverride.get(), enable);
    }
    if (state.themesRemoveOverride)
    {
        EnableWindow(state.themesRemoveOverride.get(), enable);
    }
    if (state.themesSaveTheme)
    {
        EnableWindow(state.themesSaveTheme.get(), enable);
    }
    if (state.themesDuplicateTheme)
    {
        EnableWindow(state.themesDuplicateTheme.get(), editable ? FALSE : TRUE);
    }
}

void RefreshThemesPage(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    EnsureThemesBaseComboItems(state);
    PopulateThemesThemeCombo(state);

    const auto themeIdOpt = TryGetSelectedThemeId(state);
    if (! themeIdOpt.has_value())
    {
        return;
    }

    const std::wstring_view themeId = themeIdOpt.value();

    bool editable   = false;
    const auto* def = FindThemeDefinitionForDisplay(state, themeId, editable);

    state.refreshingThemesPage = true;
    const auto reset           = wil::scope_exit([&] { state.refreshingThemesPage = false; });

    if (state.themesNote)
    {
        if (editable)
        {
            SetWindowTextW(state.themesNote.get(), L"");
        }
        else if (def)
        {
            const std::wstring note = LoadStringResource(nullptr, IDS_PREFS_THEMES_NOTE_DISK_THEME);
            SetWindowTextW(state.themesNote.get(), note.c_str());
        }
        else
        {
            const std::wstring note = LoadStringResource(nullptr, IDS_PREFS_THEMES_NOTE_BUILTIN_THEME);
            SetWindowTextW(state.themesNote.get(), note.c_str());
        }
    }

    if (state.themesNameEdit)
    {
        if (def)
        {
            SetWindowTextW(state.themesNameEdit.get(), def->name.c_str());
        }
        else
        {
            const std::wstring builtinName = GetBuiltinThemeName(themeId);
            SetWindowTextW(state.themesNameEdit.get(), builtinName.c_str());
        }
    }

    if (state.themesBaseCombo)
    {
        int select = 0;
        if (def)
        {
            for (size_t i = 0; i < kBuiltinThemeOptions.size(); ++i)
            {
                if (kBuiltinThemeOptions[i].id == def->baseThemeId)
                {
                    select = static_cast<int>(i) + 1;
                    break;
                }
            }
        }
        SendMessageW(state.themesBaseCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(select), 0);
        PrefsUi::InvalidateComboBox(state.themesBaseCombo.get());
    }

    UpdateThemesEnabled(state, editable);

    RefreshThemesColorsList(host, state, themeId, def);

    ThemesPane::UpdateEditorFromSelection(host, state);
    SendMessageW(host, WM_SIZE, 0, 0);
    InvalidateRect(host, nullptr, TRUE);
}

[[nodiscard]] ThemeMode ThemeModeFromThemeId(std::wstring_view id) noexcept
{
    if (id == L"builtin/light")
    {
        return ThemeMode::Light;
    }
    if (id == L"builtin/dark")
    {
        return ThemeMode::Dark;
    }
    if (id == L"builtin/rainbow")
    {
        return ThemeMode::Rainbow;
    }
    if (id == L"builtin/highContrast")
    {
        return ThemeMode::HighContrast;
    }
    return ThemeMode::System;
}

[[nodiscard]] float AlphaFromArgb(uint32_t argb) noexcept
{
    return static_cast<float>((argb >> 24) & 0xFFu) / 255.0f;
}

[[nodiscard]] std::optional<uint32_t> FindColorOverride(const std::unordered_map<std::wstring, uint32_t>& colors, std::wstring_view key) noexcept
{
    for (const auto& [storedKey, value] : colors)
    {
        if (storedKey == key)
        {
            return value;
        }
    }
    return std::nullopt;
}

[[nodiscard]] std::optional<D2D1::ColorF> FindAccentOverride(const std::unordered_map<std::wstring, uint32_t>& colors) noexcept
{
    const auto argb = FindColorOverride(colors, L"app.accent");
    if (! argb)
    {
        return std::nullopt;
    }

    const COLORREF rgb = ColorRefFromArgb(*argb);
    return ColorFromCOLORREF(rgb, AlphaFromArgb(*argb));
}

void ApplyDialogThemeOverrides(AppTheme& theme, const std::unordered_map<std::wstring, uint32_t>& colors) noexcept
{
    const auto applyColorRef = [&](std::wstring_view key, COLORREF& target) noexcept
    {
        const auto argb = FindColorOverride(colors, key);
        if (! argb)
        {
            return;
        }
        target = ColorRefFromArgb(*argb);
    };

    const auto applyD2D = [&](std::wstring_view key, D2D1::ColorF& target) noexcept
    {
        const auto argb = FindColorOverride(colors, key);
        if (! argb)
        {
            return;
        }
        const COLORREF rgb = ColorRefFromArgb(*argb);
        target             = ColorFromCOLORREF(rgb, AlphaFromArgb(*argb));
    };

    applyD2D(L"app.accent", theme.accent);
    applyColorRef(L"window.background", theme.windowBackground);

    applyColorRef(L"menu.background", theme.menu.background);
    applyColorRef(L"menu.text", theme.menu.text);
    applyColorRef(L"menu.disabledText", theme.menu.disabledText);
    applyColorRef(L"menu.selectionBg", theme.menu.selectionBg);
    applyColorRef(L"menu.selectionText", theme.menu.selectionText);
    applyColorRef(L"menu.separator", theme.menu.separator);
    applyColorRef(L"menu.border", theme.menu.border);
}

[[nodiscard]] uint32_t ArgbFromColorRef(COLORREF rgb, uint8_t alpha = 0xFFu) noexcept
{
    const uint32_t r = static_cast<uint32_t>(GetRValue(rgb));
    const uint32_t g = static_cast<uint32_t>(GetGValue(rgb));
    const uint32_t b = static_cast<uint32_t>(GetBValue(rgb));
    return (static_cast<uint32_t>(alpha) << 24) | (r << 16) | (g << 8) | b;
}

[[nodiscard]] uint32_t ArgbFromD2DColorF(const D2D1::ColorF& color) noexcept
{
    const auto clampByte = [](float value) noexcept -> uint32_t
    {
        value = std::clamp(value, 0.0f, 1.0f);
        return static_cast<uint32_t>(value * 255.0f + 0.5f);
    };

    const uint32_t a = clampByte(color.a);
    const uint32_t r = clampByte(color.r);
    const uint32_t g = clampByte(color.g);
    const uint32_t b = clampByte(color.b);
    return (a << 24) | (r << 16) | (g << 8) | b;
}

void ApplyAppThemeOverrides(AppTheme& theme, const std::unordered_map<std::wstring, uint32_t>& colors) noexcept
{
    const auto applyColorRef = [&](std::wstring_view key, COLORREF& target) noexcept
    {
        const auto argb = FindColorOverride(colors, key);
        if (! argb)
        {
            return;
        }
        target = ColorRefFromArgb(*argb);
    };

    const auto applyD2D = [&](std::wstring_view key, D2D1::ColorF& target) noexcept
    {
        const auto argb = FindColorOverride(colors, key);
        if (! argb)
        {
            return;
        }
        const COLORREF rgb = ColorRefFromArgb(*argb);
        target             = ColorFromCOLORREF(rgb, AlphaFromArgb(*argb));
    };

    applyD2D(L"app.accent", theme.accent);
    applyColorRef(L"window.background", theme.windowBackground);

    applyColorRef(L"menu.background", theme.menu.background);
    applyColorRef(L"menu.text", theme.menu.text);
    applyColorRef(L"menu.disabledText", theme.menu.disabledText);
    applyColorRef(L"menu.selectionBg", theme.menu.selectionBg);
    applyColorRef(L"menu.selectionText", theme.menu.selectionText);
    applyColorRef(L"menu.separator", theme.menu.separator);
    applyColorRef(L"menu.border", theme.menu.border);

    applyD2D(L"navigation.background", theme.navigationView.background);
    applyD2D(L"navigation.backgroundHover", theme.navigationView.backgroundHover);
    applyD2D(L"navigation.backgroundPressed", theme.navigationView.backgroundPressed);
    applyD2D(L"navigation.text", theme.navigationView.text);
    applyD2D(L"navigation.separator", theme.navigationView.separator);
    applyD2D(L"navigation.accent", theme.navigationView.accent);
    applyD2D(L"navigation.progressOk", theme.navigationView.progressOk);
    applyD2D(L"navigation.progressWarn", theme.navigationView.progressWarn);
    applyD2D(L"navigation.progressBackground", theme.navigationView.progressBackground);

    if (const auto argb = FindColorOverride(colors, L"navigation.background"))
    {
        const COLORREF rgb                 = ColorRefFromArgb(*argb);
        theme.navigationView.gdiBackground = rgb;
        theme.navigationView.gdiBorder     = rgb;
    }

    if (const auto argb = FindColorOverride(colors, L"navigation.separator"))
    {
        theme.navigationView.gdiBorderPen = ColorRefFromArgb(*argb);
    }

    applyD2D(L"folderView.background", theme.folderView.backgroundColor);
    applyD2D(L"folderView.itemBackgroundNormal", theme.folderView.itemBackgroundNormal);
    applyD2D(L"folderView.itemBackgroundHovered", theme.folderView.itemBackgroundHovered);
    applyD2D(L"folderView.itemBackgroundSelected", theme.folderView.itemBackgroundSelected);
    applyD2D(L"folderView.itemBackgroundSelectedInactive", theme.folderView.itemBackgroundSelectedInactive);
    applyD2D(L"folderView.itemBackgroundFocused", theme.folderView.itemBackgroundFocused);
    applyD2D(L"folderView.textNormal", theme.folderView.textNormal);
    applyD2D(L"folderView.textSelected", theme.folderView.textSelected);
    applyD2D(L"folderView.textSelectedInactive", theme.folderView.textSelectedInactive);
    applyD2D(L"folderView.textDisabled", theme.folderView.textDisabled);
    applyD2D(L"folderView.focusBorder", theme.folderView.focusBorder);
    applyD2D(L"folderView.gridLines", theme.folderView.gridLines);
    applyD2D(L"folderView.errorBackground", theme.folderView.errorBackground);
    applyD2D(L"folderView.errorText", theme.folderView.errorText);
    applyD2D(L"folderView.warningBackground", theme.folderView.warningBackground);
    applyD2D(L"folderView.warningText", theme.folderView.warningText);
    applyD2D(L"folderView.infoBackground", theme.folderView.infoBackground);
    applyD2D(L"folderView.infoText", theme.folderView.infoText);

    theme.fileOperations.progressBackground = theme.navigationView.progressBackground;
    theme.fileOperations.progressTotal      = theme.navigationView.progressOk;
    theme.fileOperations.progressItem       = theme.navigationView.accent;

    const D2D1::ColorF menuBorder   = ColorFromCOLORREF(theme.menu.border);
    const D2D1::ColorF menuDisabled = ColorFromCOLORREF(theme.menu.disabledText);

    theme.fileOperations.graphBackground =
        D2D1::ColorF(theme.fileOperations.progressBackground.r, theme.fileOperations.progressBackground.g, theme.fileOperations.progressBackground.b, 0.35f);
    theme.fileOperations.graphGrid      = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.35f);
    theme.fileOperations.graphLimit     = D2D1::ColorF(menuDisabled.r, menuDisabled.g, menuDisabled.b, 0.85f);
    theme.fileOperations.graphLine      = theme.fileOperations.progressItem;
    theme.fileOperations.scrollbarTrack = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.12f);
    theme.fileOperations.scrollbarThumb = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.40f);

    applyD2D(L"fileOps.progressBackground", theme.fileOperations.progressBackground);
    applyD2D(L"fileOps.progressTotal", theme.fileOperations.progressTotal);
    applyD2D(L"fileOps.progressItem", theme.fileOperations.progressItem);
    applyD2D(L"fileOps.graphBackground", theme.fileOperations.graphBackground);
    applyD2D(L"fileOps.graphGrid", theme.fileOperations.graphGrid);
    applyD2D(L"fileOps.graphLimit", theme.fileOperations.graphLimit);
    applyD2D(L"fileOps.graphLine", theme.fileOperations.graphLine);
    applyD2D(L"fileOps.scrollbarTrack", theme.fileOperations.scrollbarTrack);
    applyD2D(L"fileOps.scrollbarThumb", theme.fileOperations.scrollbarThumb);

    if (! FindColorOverride(colors, L"folderView.itemBackgroundSelectedInactive"))
    {
        if (const auto argb = FindColorOverride(colors, L"folderView.itemBackgroundSelected"))
        {
            const float inactiveSelectionAlphaScale = theme.highContrast ? 0.80f : 0.65f;
            const COLORREF rgb                      = ColorRefFromArgb(*argb);
            theme.folderView.itemBackgroundSelectedInactive =
                ColorFromCOLORREF(rgb, std::clamp(AlphaFromArgb(*argb) * inactiveSelectionAlphaScale, 0.0f, 1.0f));
        }
    }

    if (! FindColorOverride(colors, L"folderView.textSelectedInactive") && ! theme.highContrast)
    {
        const float alpha             = std::clamp(theme.folderView.itemBackgroundSelectedInactive.a, 0.0f, 1.0f);
        const D2D1::ColorF background = theme.folderView.backgroundColor;
        const D2D1::ColorF overlay    = theme.folderView.itemBackgroundSelectedInactive;

        const D2D1::ColorF composite = D2D1::ColorF(overlay.r * alpha + background.r * (1.0f - alpha),
                                                    overlay.g * alpha + background.g * (1.0f - alpha),
                                                    overlay.b * alpha + background.b * (1.0f - alpha),
                                                    1.0f);

        const COLORREF contrastText           = ChooseContrastingTextColor(ColorToCOLORREF(composite));
        theme.folderView.textSelectedInactive = ColorFromCOLORREF(contrastText);
    }
}

[[nodiscard]] MonitorTextViewTheme ResolveMonitorThemeForDisplay(std::wstring_view baseThemeId,
                                                                 const std::unordered_map<std::wstring, uint32_t>* overrides) noexcept
{
    const ThemeMode mode = ThemeModeFromThemeId(baseThemeId);
    MonitorTextViewTheme theme{};

    if (mode == ThemeMode::Dark)
    {
        theme.bg              = D2D1::ColorF(0.08f, 0.08f, 0.08f);
        theme.fg              = D2D1::ColorF(0.90f, 0.90f, 0.90f);
        theme.caret           = D2D1::ColorF(0.90f, 0.90f, 0.90f);
        theme.selection       = D2D1::ColorF(0.20f, 0.55f, 0.95f, 0.35f);
        theme.searchHighlight = D2D1::ColorF(1.00f, 0.85f, 0.05f, 0.35f);
        theme.gutterBg        = D2D1::ColorF(0.12f, 0.12f, 0.12f);
        theme.gutterFg        = D2D1::ColorF(0.65f, 0.65f, 0.65f);
        theme.metaText        = D2D1::ColorF(0.65f, 0.65f, 0.65f);
        theme.metaError       = D2D1::ColorF(1.00f, 0.35f, 0.35f);
        theme.metaWarning     = D2D1::ColorF(1.00f, 0.70f, 0.25f);
        theme.metaInfo        = D2D1::ColorF(0.40f, 0.70f, 1.00f);
        theme.metaDebug       = D2D1::ColorF(0.75f, 0.55f, 1.00f);
    }
    else if (mode == ThemeMode::Rainbow)
    {
        theme.bg              = D2D1::ColorF(0.10f, 0.10f, 0.10f);
        theme.fg              = D2D1::ColorF(0.95f, 0.95f, 0.95f);
        theme.caret           = D2D1::ColorF(0.95f, 0.95f, 0.95f);
        theme.selection       = D2D1::ColorF(0.35f, 0.75f, 1.00f, 0.35f);
        theme.searchHighlight = D2D1::ColorF(1.00f, 0.85f, 0.05f, 0.40f);
        theme.gutterBg        = D2D1::ColorF(0.15f, 0.15f, 0.15f);
        theme.gutterFg        = D2D1::ColorF(0.70f, 0.70f, 0.70f);
        theme.metaText        = D2D1::ColorF(0.70f, 0.70f, 0.70f);
        theme.metaError       = D2D1::ColorF(1.00f, 0.45f, 0.45f);
        theme.metaWarning     = D2D1::ColorF(1.00f, 0.75f, 0.30f);
        theme.metaInfo        = D2D1::ColorF(0.50f, 0.80f, 1.00f);
        theme.metaDebug       = D2D1::ColorF(0.80f, 0.60f, 1.00f);
    }
    else if (mode == ThemeMode::HighContrast)
    {
        const COLORREF window = GetSysColor(COLOR_WINDOW);
        const COLORREF text   = GetSysColor(COLOR_WINDOWTEXT);
        const COLORREF sel    = GetSysColor(COLOR_HIGHLIGHT);
        theme.bg              = ColorFromCOLORREF(window, 1.0f);
        theme.fg              = ColorFromCOLORREF(text, 1.0f);
        theme.caret           = ColorFromCOLORREF(text, 1.0f);
        theme.selection       = ColorFromCOLORREF(sel, 0.40f);
        theme.searchHighlight = D2D1::ColorF(1.00f, 0.85f, 0.05f, 0.50f);
        theme.gutterBg        = ColorFromCOLORREF(window, 1.0f);
        theme.gutterFg        = ColorFromCOLORREF(text, 1.0f);
        theme.metaText        = ColorFromCOLORREF(text, 1.0f);
        theme.metaError       = ColorFromCOLORREF(text, 1.0f);
        theme.metaWarning     = ColorFromCOLORREF(text, 1.0f);
        theme.metaInfo        = ColorFromCOLORREF(text, 1.0f);
        theme.metaDebug       = ColorFromCOLORREF(text, 1.0f);
    }

    if (overrides)
    {
        const auto applyOverride = [&](std::wstring_view key, D2D1::ColorF& target) noexcept
        {
            const auto argb = FindColorOverride(*overrides, key);
            if (! argb)
            {
                return;
            }
            const COLORREF rgb = ColorRefFromArgb(*argb);
            target             = ColorFromCOLORREF(rgb, AlphaFromArgb(*argb));
        };

        applyOverride(L"monitor.textView.bg", theme.bg);
        applyOverride(L"monitor.textView.fg", theme.fg);
        applyOverride(L"monitor.textView.caret", theme.caret);
        applyOverride(L"monitor.textView.selection", theme.selection);
        applyOverride(L"monitor.textView.searchHighlight", theme.searchHighlight);
        applyOverride(L"monitor.textView.gutterBg", theme.gutterBg);
        applyOverride(L"monitor.textView.gutterFg", theme.gutterFg);
        applyOverride(L"monitor.textView.metaText", theme.metaText);
        applyOverride(L"monitor.textView.metaError", theme.metaError);
        applyOverride(L"monitor.textView.metaWarning", theme.metaWarning);
        applyOverride(L"monitor.textView.metaInfo", theme.metaInfo);
        applyOverride(L"monitor.textView.metaDebug", theme.metaDebug);
    }

    return theme;
}

[[nodiscard]] std::optional<uint32_t> TryGetEffectiveThemeColorArgb(const AppTheme& appTheme,
                                                                    const MonitorTextViewTheme& monitorTheme,
                                                                    const std::unordered_map<std::wstring, uint32_t>* overrides,
                                                                    std::wstring_view key) noexcept
{
    if (key == L"app.accent")
    {
        return ArgbFromD2DColorF(appTheme.accent);
    }
    if (key == L"window.background")
    {
        return ArgbFromColorRef(appTheme.windowBackground);
    }

    if (key == L"menu.background")
    {
        return ArgbFromColorRef(appTheme.menu.background);
    }
    if (key == L"menu.text")
    {
        return ArgbFromColorRef(appTheme.menu.text);
    }
    if (key == L"menu.disabledText")
    {
        return ArgbFromColorRef(appTheme.menu.disabledText);
    }
    if (key == L"menu.selectionBg")
    {
        return ArgbFromColorRef(appTheme.menu.selectionBg);
    }
    if (key == L"menu.selectionText")
    {
        return ArgbFromColorRef(appTheme.menu.selectionText);
    }
    if (key == L"menu.separator")
    {
        return ArgbFromColorRef(appTheme.menu.separator);
    }
    if (key == L"menu.border")
    {
        return ArgbFromColorRef(appTheme.menu.border);
    }

    if (key == L"navigation.background")
    {
        return ArgbFromD2DColorF(appTheme.navigationView.background);
    }
    if (key == L"navigation.backgroundHover")
    {
        return ArgbFromD2DColorF(appTheme.navigationView.backgroundHover);
    }
    if (key == L"navigation.backgroundPressed")
    {
        return ArgbFromD2DColorF(appTheme.navigationView.backgroundPressed);
    }
    if (key == L"navigation.text")
    {
        return ArgbFromD2DColorF(appTheme.navigationView.text);
    }
    if (key == L"navigation.separator")
    {
        return ArgbFromD2DColorF(appTheme.navigationView.separator);
    }
    if (key == L"navigation.accent")
    {
        return ArgbFromD2DColorF(appTheme.navigationView.accent);
    }
    if (key == L"navigation.progressOk")
    {
        return ArgbFromD2DColorF(appTheme.navigationView.progressOk);
    }
    if (key == L"navigation.progressWarn")
    {
        return ArgbFromD2DColorF(appTheme.navigationView.progressWarn);
    }
    if (key == L"navigation.progressBackground")
    {
        return ArgbFromD2DColorF(appTheme.navigationView.progressBackground);
    }

    if (key == L"folderView.background")
    {
        return ArgbFromD2DColorF(appTheme.folderView.backgroundColor);
    }
    if (key == L"folderView.itemBackgroundNormal")
    {
        return ArgbFromD2DColorF(appTheme.folderView.itemBackgroundNormal);
    }
    if (key == L"folderView.itemBackgroundHovered")
    {
        return ArgbFromD2DColorF(appTheme.folderView.itemBackgroundHovered);
    }
    if (key == L"folderView.itemBackgroundSelected")
    {
        return ArgbFromD2DColorF(appTheme.folderView.itemBackgroundSelected);
    }
    if (key == L"folderView.itemBackgroundSelectedInactive")
    {
        return ArgbFromD2DColorF(appTheme.folderView.itemBackgroundSelectedInactive);
    }
    if (key == L"folderView.itemBackgroundFocused")
    {
        return ArgbFromD2DColorF(appTheme.folderView.itemBackgroundFocused);
    }
    if (key == L"folderView.textNormal")
    {
        return ArgbFromD2DColorF(appTheme.folderView.textNormal);
    }
    if (key == L"folderView.textSelected")
    {
        return ArgbFromD2DColorF(appTheme.folderView.textSelected);
    }
    if (key == L"folderView.textSelectedInactive")
    {
        return ArgbFromD2DColorF(appTheme.folderView.textSelectedInactive);
    }
    if (key == L"folderView.textDisabled")
    {
        return ArgbFromD2DColorF(appTheme.folderView.textDisabled);
    }
    if (key == L"folderView.focusBorder")
    {
        return ArgbFromD2DColorF(appTheme.folderView.focusBorder);
    }
    if (key == L"folderView.gridLines")
    {
        return ArgbFromD2DColorF(appTheme.folderView.gridLines);
    }
    if (key == L"folderView.errorBackground")
    {
        return ArgbFromD2DColorF(appTheme.folderView.errorBackground);
    }
    if (key == L"folderView.errorText")
    {
        return ArgbFromD2DColorF(appTheme.folderView.errorText);
    }
    if (key == L"folderView.warningBackground")
    {
        return ArgbFromD2DColorF(appTheme.folderView.warningBackground);
    }
    if (key == L"folderView.warningText")
    {
        return ArgbFromD2DColorF(appTheme.folderView.warningText);
    }
    if (key == L"folderView.infoBackground")
    {
        return ArgbFromD2DColorF(appTheme.folderView.infoBackground);
    }
    if (key == L"folderView.infoText")
    {
        return ArgbFromD2DColorF(appTheme.folderView.infoText);
    }

    if (key == L"monitor.textView.bg")
    {
        return ArgbFromD2DColorF(monitorTheme.bg);
    }
    if (key == L"monitor.textView.fg")
    {
        return ArgbFromD2DColorF(monitorTheme.fg);
    }
    if (key == L"monitor.textView.caret")
    {
        return ArgbFromD2DColorF(monitorTheme.caret);
    }
    if (key == L"monitor.textView.selection")
    {
        return ArgbFromD2DColorF(monitorTheme.selection);
    }
    if (key == L"monitor.textView.searchHighlight")
    {
        return ArgbFromD2DColorF(monitorTheme.searchHighlight);
    }
    if (key == L"monitor.textView.gutterBg")
    {
        return ArgbFromD2DColorF(monitorTheme.gutterBg);
    }
    if (key == L"monitor.textView.gutterFg")
    {
        return ArgbFromD2DColorF(monitorTheme.gutterFg);
    }
    if (key == L"monitor.textView.metaText")
    {
        return ArgbFromD2DColorF(monitorTheme.metaText);
    }
    if (key == L"monitor.textView.metaError")
    {
        return ArgbFromD2DColorF(monitorTheme.metaError);
    }
    if (key == L"monitor.textView.metaWarning")
    {
        return ArgbFromD2DColorF(monitorTheme.metaWarning);
    }
    if (key == L"monitor.textView.metaInfo")
    {
        return ArgbFromD2DColorF(monitorTheme.metaInfo);
    }
    if (key == L"monitor.textView.metaDebug")
    {
        return ArgbFromD2DColorF(monitorTheme.metaDebug);
    }

    if (key == L"fileOps.progressBackground")
    {
        return ArgbFromD2DColorF(appTheme.fileOperations.progressBackground);
    }
    if (key == L"fileOps.progressTotal")
    {
        return ArgbFromD2DColorF(appTheme.fileOperations.progressTotal);
    }
    if (key == L"fileOps.progressItem")
    {
        return ArgbFromD2DColorF(appTheme.fileOperations.progressItem);
    }
    if (key == L"fileOps.graphBackground")
    {
        return ArgbFromD2DColorF(appTheme.fileOperations.graphBackground);
    }
    if (key == L"fileOps.graphGrid")
    {
        return ArgbFromD2DColorF(appTheme.fileOperations.graphGrid);
    }
    if (key == L"fileOps.graphLimit")
    {
        return ArgbFromD2DColorF(appTheme.fileOperations.graphLimit);
    }
    if (key == L"fileOps.graphLine")
    {
        return ArgbFromD2DColorF(appTheme.fileOperations.graphLine);
    }
    if (key == L"fileOps.scrollbarTrack")
    {
        return ArgbFromD2DColorF(appTheme.fileOperations.scrollbarTrack);
    }
    if (key == L"fileOps.scrollbarThumb")
    {
        return ArgbFromD2DColorF(appTheme.fileOperations.scrollbarThumb);
    }

    if (overrides)
    {
        return FindColorOverride(*overrides, key);
    }

    return std::nullopt;
}

void BeginNewThemeCreation(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    EnsureThemeFileThemesLoaded(state);

    const std::wstring defaultName    = LoadStringResource(nullptr, IDS_PREFS_THEMES_DEFAULT_NEW_NAME);
    std::wstring_view suggestedBaseId = L"builtin/system";
    if (IsBuiltinThemeId(state.workingSettings.theme.currentThemeId))
    {
        suggestedBaseId = state.workingSettings.theme.currentThemeId;
    }
    else
    {
        bool editable = false;
        if (const auto* existing = FindThemeDefinitionForDisplay(state, state.workingSettings.theme.currentThemeId, editable); existing)
        {
            if (! existing->baseThemeId.empty())
            {
                suggestedBaseId = existing->baseThemeId;
            }
        }
    }

    Common::Settings::ThemeDefinition def;
    def.id          = MakeUniqueUserThemeId(state, defaultName);
    def.name        = defaultName;
    def.baseThemeId = std::wstring(suggestedBaseId);

    state.workingSettings.theme.themes.push_back(std::move(def));
    state.workingSettings.theme.currentThemeId = state.workingSettings.theme.themes.back().id;

    SetDirty(dlg, state);
    RefreshThemesPage(host, state);

    if (state.themesNameEdit)
    {
        SetFocus(state.themesNameEdit.get());
        SendMessageW(state.themesNameEdit.get(), EM_SETSEL, 0, -1);
    }
}

void DuplicateSelectedTheme(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    const auto themeIdOpt = TryGetSelectedThemeId(state);
    if (! themeIdOpt.has_value())
    {
        return;
    }

    const std::wstring_view themeId = themeIdOpt.value();

    EnsureThemeFileThemesLoaded(state);

    bool editable         = false;
    const auto* sourceDef = FindThemeDefinitionForDisplay(state, themeId, editable);
    if (editable)
    {
        return;
    }

    std::wstring sourceNameText;
    std::wstring_view sourceName;
    if (const auto* comboItem = TryGetSelectedThemeComboItem(state))
    {
        sourceName = comboItem->displayName;
    }
    if (sourceName.empty())
    {
        if (sourceDef && ! sourceDef->name.empty())
        {
            sourceName = sourceDef->name;
        }
        else if (sourceDef)
        {
            sourceName = sourceDef->id;
        }
        else
        {
            sourceNameText = GetBuiltinThemeName(themeId);
            if (sourceNameText.empty())
            {
                sourceNameText = LoadStringResource(nullptr, IDS_PREFS_THEMES_DEFAULT_NEW_NAME);
            }
            sourceName = sourceNameText;
        }
    }

    std::wstring newName;
    newName = FormatStringResource(nullptr, IDS_PREFS_THEMES_DUPLICATE_NAME_FMT, sourceName);

    if (newName.empty())
    {
        newName = LoadStringResource(nullptr, IDS_PREFS_THEMES_DEFAULT_NEW_NAME);
    }
    if (newName.size() > 64u)
    {
        newName.resize(64u);
    }

    Common::Settings::ThemeDefinition def;
    def.id   = MakeUniqueUserThemeId(state, newName);
    def.name = newName;

    if (sourceDef)
    {
        def.baseThemeId = sourceDef->baseThemeId.empty() ? std::wstring(themeId) : sourceDef->baseThemeId;
        def.colors      = sourceDef->colors;
    }
    else
    {
        def.baseThemeId = std::wstring(themeId);
    }

    state.workingSettings.theme.themes.push_back(std::move(def));
    state.workingSettings.theme.currentThemeId = state.workingSettings.theme.themes.back().id;

    SetDirty(dlg, state);
    RefreshThemesPage(host, state);

    if (state.themesNameEdit)
    {
        SetFocus(state.themesNameEdit.get());
        SendMessageW(state.themesNameEdit.get(), EM_SETSEL, 0, -1);
    }
}

void SyncSelectedUserThemeIdToName(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    const auto themeIdOpt = TryGetSelectedThemeId(state);
    if (! themeIdOpt.has_value())
    {
        return;
    }

    auto* def = FindWorkingThemeDefinition(state, themeIdOpt.value());
    if (! def)
    {
        return;
    }

    if (def->id.rfind(L"user/", 0) != 0 || def->name.empty())
    {
        return;
    }

    const std::wstring oldId = def->id;
    const std::wstring newId = MakeUniqueUserThemeIdForRename(state, def->name, oldId);
    if (newId.empty() || newId == oldId)
    {
        return;
    }

    def->id = newId;
    if (state.workingSettings.theme.currentThemeId == oldId)
    {
        state.workingSettings.theme.currentThemeId = newId;
    }

    SetDirty(dlg, state);
    RefreshThemesPage(host, state);
}

void ApplyThemeTemporarily(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg || ! state.settings)
    {
        return;
    }

    Common::Settings::Settings preview = *state.settings;
    preview.theme                      = state.workingSettings.theme;
    *state.settings                    = std::move(preview);

    state.previewApplied = true;
    ApplyThemeToPreferencesDialog(dlg, state, ResolveThemeFromSettingsForDialog(*state.settings));
    if (state.pageHost)
    {
        RedrawWindow(state.pageHost, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
    }
    RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN | RDW_UPDATENOW);
    if (state.owner)
    {
        PostMessageW(state.owner, WndMsg::kSettingsApplied, 0, 0);
    }
}

void PickThemeColorIntoEditor(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg || ! state.themesColorEdit)
    {
        return;
    }

    uint32_t currentArgb           = 0xFF000000u;
    uint32_t alpha                 = 0xFFu;
    const std::wstring currentText = PrefsUi::GetWindowTextString(state.themesColorEdit.get());
    if (! currentText.empty() && Common::Settings::TryParseColor(currentText, currentArgb))
    {
        alpha = (currentArgb >> 24) & 0xFFu;
    }
    else
    {
        alpha       = 0xFFu;
        currentArgb = (alpha << 24) | (static_cast<uint32_t>(GetRValue(state.theme.menu.selectionBg)) << 16) |
                      (static_cast<uint32_t>(GetGValue(state.theme.menu.selectionBg)) << 8) | static_cast<uint32_t>(GetBValue(state.theme.menu.selectionBg));
    }

    static COLORREF customColors[16] = {};

    CHOOSECOLORW cc{};
    cc.lStructSize  = sizeof(cc);
    cc.hwndOwner    = dlg;
    cc.rgbResult    = RGB((currentArgb >> 16) & 0xFFu, (currentArgb >> 8) & 0xFFu, currentArgb & 0xFFu);
    cc.lpCustColors = customColors;
    cc.Flags        = CC_FULLOPEN | CC_RGBINIT;

    if (! ChooseColorW(&cc))
    {
        return;
    }

    const uint32_t rgb = (static_cast<uint32_t>(GetRValue(cc.rgbResult)) << 16) | (static_cast<uint32_t>(GetGValue(cc.rgbResult)) << 8) |
                         static_cast<uint32_t>(GetBValue(cc.rgbResult));
    const uint32_t argb = (alpha << 24) | rgb;

    std::wstring text;
    text                       = Common::Settings::FormatColor(argb);
    state.refreshingThemesPage = true;
    const auto reset           = wil::scope_exit([&] { state.refreshingThemesPage = false; });
    SetWindowTextW(state.themesColorEdit.get(), text.c_str());
    if (state.themesColorSwatch)
    {
        InvalidateRect(state.themesColorSwatch.get(), nullptr, TRUE);
    }
}

void SetThemeOverrideFromEditor(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    const auto themeIdOpt = TryGetSelectedThemeId(state);
    if (! themeIdOpt.has_value())
    {
        return;
    }

    auto* def = FindWorkingThemeDefinition(state, themeIdOpt.value());
    if (! def)
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_WARNING, LoadStringResource(nullptr, IDS_CAPTION_WARNING), LoadStringResource(nullptr, IDS_PREFS_THEMES_WARNING_SELECT_USER_EDIT));
        return;
    }

    std::wstring key = PrefsUi::GetWindowTextString(state.themesKeyEdit.get());
    key.erase(key.begin(), std::find_if(key.begin(), key.end(), [](wchar_t ch) noexcept { return ! std::iswspace(ch); }));
    key.erase(std::find_if(key.rbegin(), key.rend(), [](wchar_t ch) noexcept { return ! std::iswspace(ch); }).base(), key.end());
    if (! IsValidThemeColorKey(key))
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_WARNING, LoadStringResource(nullptr, IDS_CAPTION_WARNING), LoadStringResource(nullptr, IDS_PREFS_THEMES_WARNING_ENTER_COLOR_KEY));
        return;
    }

    const std::wstring valueText = PrefsUi::GetWindowTextString(state.themesColorEdit.get());
    uint32_t argb                = 0;
    if (valueText.empty() || ! Common::Settings::TryParseColor(valueText, argb))
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_WARNING, LoadStringResource(nullptr, IDS_CAPTION_WARNING), LoadStringResource(nullptr, IDS_PREFS_THEMES_WARNING_ENTER_COLOR_VALUE));
        return;
    }

    def->colors[key] = argb;

    SetDirty(dlg, state);
    RefreshThemesPage(host, state);
}

void ClearThemeOverrideFromEditor(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    const auto themeIdOpt = TryGetSelectedThemeId(state);
    if (! themeIdOpt.has_value())
    {
        return;
    }

    auto* def = FindWorkingThemeDefinition(state, themeIdOpt.value());
    if (! def)
    {
        return;
    }

    std::wstring key = PrefsUi::GetWindowTextString(state.themesKeyEdit.get());
    key.erase(key.begin(), std::find_if(key.begin(), key.end(), [](wchar_t ch) noexcept { return ! std::iswspace(ch); }));
    key.erase(std::find_if(key.rbegin(), key.rend(), [](wchar_t ch) noexcept { return ! std::iswspace(ch); }).base(), key.end());
    if (key.empty())
    {
        return;
    }

    const size_t removed = def->colors.erase(key);
    if (removed == 0)
    {
        return;
    }

    SetDirty(dlg, state);
    RefreshThemesPage(host, state);
}

void LoadThemeFromFile(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    std::filesystem::path path;
    if (! TryBrowseThemeFile(dlg, false, {}, path))
    {
        return;
    }

    std::string jsonText;
    if (! PrefsFile::TryReadFileToString(path, jsonText))
    {
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_PREFS_THEMES_ERROR_READ_FILE));
        return;
    }

    Common::Settings::ThemeDefinition imported;
    std::wstring error;
    if (! ParseThemeDefinitionJson(jsonText, imported, error))
    {
        if (error.empty())
        {
            error = LoadStringResource(nullptr, IDS_PREFS_THEMES_ERROR_LOAD_FILE);
        }
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), error);
        return;
    }

    auto& themes = state.workingSettings.theme.themes;
    auto it      = std::find_if(themes.begin(), themes.end(), [&](const auto& t) noexcept { return t.id == imported.id; });
    if (it != themes.end())
    {
        *it                                        = std::move(imported);
        state.workingSettings.theme.currentThemeId = it->id;
    }
    else
    {
        themes.push_back(std::move(imported));
        state.workingSettings.theme.currentThemeId = themes.back().id;
    }

    SetDirty(dlg, state);
    RefreshThemesPage(host, state);
}

void SaveThemeToFile(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    const auto themeIdOpt = TryGetSelectedThemeId(state);
    if (! themeIdOpt.has_value())
    {
        return;
    }

    auto* def = FindWorkingThemeDefinition(state, themeIdOpt.value());
    if (! def)
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_WARNING, LoadStringResource(nullptr, IDS_CAPTION_WARNING), LoadStringResource(nullptr, IDS_PREFS_THEMES_WARNING_SELECT_USER_SAVE));
        return;
    }

    const std::wstring suggested = MakeSuggestedThemeFileName(def->id, def->name);
    std::filesystem::path path;
    if (! TryBrowseThemeFile(dlg, true, suggested, path))
    {
        return;
    }

    std::string json;
    if (! BuildThemeDefinitionExportJson(*def, json))
    {
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_PREFS_THEMES_ERROR_BUILD_FILE));
        return;
    }

    if (! PrefsFile::TryWriteFileFromString(path, json))
    {
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_PREFS_THEMES_ERROR_WRITE_FILE));
        return;
    }
}
} // namespace

[[nodiscard]] AppTheme ResolveThemeFromSettingsForDialog(const Common::Settings::Settings& settings) noexcept
{
    std::wstring_view themeId                       = settings.theme.currentThemeId;
    const Common::Settings::ThemeDefinition* custom = nullptr;
    if (themeId.rfind(L"user/", 0) == 0)
    {
        custom = FindThemeDefinitionById(settings.theme.themes, themeId);
    }

    ThemeMode baseMode = ThemeModeFromThemeId(themeId);
    std::optional<D2D1::ColorF> accentOverride;
    const std::unordered_map<std::wstring, uint32_t>* overrides = nullptr;
    if (custom)
    {
        baseMode       = ThemeModeFromThemeId(custom->baseThemeId);
        accentOverride = FindAccentOverride(custom->colors);
        overrides      = &custom->colors;
    }

    AppTheme theme = ResolveAppTheme(baseMode, L"RedSalamander", accentOverride);
    if (overrides)
    {
        ApplyDialogThemeOverrides(theme, *overrides);
    }

    return theme;
}

void ApplyThemeToPreferencesDialog(HWND dlg, PreferencesDialogState& state, const AppTheme& theme) noexcept
{
    state.theme = theme;
    ApplyTitleBarTheme(dlg, state.theme, GetActiveWindow() == dlg);

    state.backgroundBrush.reset(CreateSolidBrush(state.theme.windowBackground));
    state.cardBackgroundColor = ThemedControls::GetControlSurfaceColor(state.theme);
    state.inputBrush.reset();
    state.inputFocusedBrush.reset();
    state.inputDisabledBrush.reset();
    state.cardBrush.reset();

    state.inputBackgroundColor         = ThemedControls::BlendColor(state.cardBackgroundColor, state.theme.windowBackground, state.theme.dark ? 50 : 30, 255);
    state.inputFocusedBackgroundColor  = ThemedControls::BlendColor(state.inputBackgroundColor, state.theme.menu.text, state.theme.dark ? 20 : 16, 255);
    state.inputDisabledBackgroundColor = ThemedControls::BlendColor(state.theme.windowBackground, state.inputBackgroundColor, state.theme.dark ? 70 : 40, 255);
    if (! state.theme.systemHighContrast)
    {
        state.cardBrush.reset(CreateSolidBrush(state.cardBackgroundColor));
        state.inputBrush.reset(CreateSolidBrush(state.inputBackgroundColor));
        state.inputFocusedBrush.reset(CreateSolidBrush(state.inputFocusedBackgroundColor));
        state.inputDisabledBrush.reset(CreateSolidBrush(state.inputDisabledBackgroundColor));
    }

    if (state.keyboardScopeCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.keyboardScopeCombo, state.theme);
    }
    if (state.panesLeftDisplayCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.panesLeftDisplayCombo, state.theme);
    }
    if (state.panesLeftSortByCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.panesLeftSortByCombo, state.theme);
    }
    if (state.panesLeftSortDirCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.panesLeftSortDirCombo, state.theme);
    }
    if (state.panesRightDisplayCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.panesRightDisplayCombo, state.theme);
    }
    if (state.panesRightSortByCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.panesRightSortByCombo, state.theme);
    }
    if (state.panesRightSortDirCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.panesRightSortDirCombo, state.theme);
    }
    if (state.viewersViewerCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.viewersViewerCombo, state.theme);
    }
    if (state.viewersList)
    {
        ThemedControls::ApplyThemeToListView(state.viewersList, state.theme);
    }
    if (state.keyboardList)
    {
        ThemedControls::ApplyThemeToListView(state.keyboardList, state.theme);
    }
    if (state.themesThemeCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.themesThemeCombo.get(), state.theme);
    }
    if (state.themesBaseCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.themesBaseCombo.get(), state.theme);
    }
    if (state.advancedMonitorFilterPresetCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.advancedMonitorFilterPresetCombo, state.theme);
    }
    if (state.themesColorsList)
    {
        ThemedControls::ApplyThemeToListView(state.themesColorsList.get(), state.theme);
    }
    if (state.pluginsList)
    {
        ThemedControls::ApplyThemeToListView(state.pluginsList, state.theme);
    }

    if (state.categoryTree)
    {
        if (state.theme.systemHighContrast)
        {
            SetWindowTheme(state.categoryTree, L"", nullptr);
            TreeView_SetBkColor(state.categoryTree, GetSysColor(COLOR_WINDOW));
            TreeView_SetTextColor(state.categoryTree, GetSysColor(COLOR_WINDOWTEXT));
        }
        else
        {
            const wchar_t* listTheme = state.theme.dark ? L"DarkMode_Explorer" : L"Explorer";
            SetWindowTheme(state.categoryTree, listTheme, nullptr);
            TreeView_SetBkColor(state.categoryTree, state.theme.windowBackground);
            TreeView_SetTextColor(state.categoryTree, state.theme.menu.text);
        }
        SendMessageW(state.categoryTree, WM_THEMECHANGED, 0, 0);
        InvalidateRect(state.categoryTree, nullptr, TRUE);
    }
    if (state.pageHost)
    {
        if (state.theme.systemHighContrast)
        {
            SetWindowTheme(state.pageHost, L"", nullptr);
        }
        else
        {
            const wchar_t* hostTheme = state.theme.dark ? L"DarkMode_Explorer" : L"Explorer";
            SetWindowTheme(state.pageHost, hostTheme, nullptr);
        }
        SendMessageW(state.pageHost, WM_THEMECHANGED, 0, 0);
        InvalidateRect(state.pageHost, nullptr, TRUE);
    }

    RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_FRAME | RDW_ERASE | RDW_ALLCHILDREN);
}

void UpdateThemesColorsListColumnWidths(HWND list, UINT dpi) noexcept
{
    if (! list)
    {
        return;
    }

    RECT rc{};
    GetClientRect(list, &rc);
    const int totalWidth = std::max(0l, rc.right - rc.left);
    if (totalWidth <= 0)
    {
        return;
    }

    const int swatchWidth = std::min(totalWidth, ThemedControls::ScaleDip(dpi, 44));
    const int minValueW   = std::min(std::max(0, totalWidth - swatchWidth), ThemedControls::ScaleDip(dpi, 110));
    const int minKeyW     = std::min(std::max(0, totalWidth - swatchWidth), ThemedControls::ScaleDip(dpi, 180));

    int keyW   = std::max(0, totalWidth - swatchWidth - minValueW);
    int valueW = std::max(0, totalWidth - swatchWidth - keyW);
    if (keyW < minKeyW)
    {
        keyW   = minKeyW;
        valueW = std::max(0, totalWidth - swatchWidth - keyW);
    }

    ListView_SetColumnWidth(list, 0, keyW);
    ListView_SetColumnWidth(list, 1, valueW);
    ListView_SetColumnWidth(list, 2, swatchWidth);
}

bool ThemesPane::EnsureCreated(HWND pageHost) noexcept
{
    return PrefsPaneHost::EnsureCreated(pageHost, _hWnd);
}

void ThemesPane::ResizeToHostClient(HWND pageHost) noexcept
{
    PrefsPaneHost::ResizeToHostClient(pageHost, _hWnd.get());
}

void ThemesPane::Show(bool visible) noexcept
{
    PrefsPaneHost::Show(_hWnd.get(), visible);
}

void ThemesPane::CreateControls(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    const std::wstring themeLabelText  = LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_THEME);
    const std::wstring nameLabelText   = LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_NAME);
    const std::wstring baseLabelText   = LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_BASE);
    const std::wstring searchLabelText = LoadStringResource(nullptr, IDS_PREFS_COMMON_SEARCH);
    const std::wstring keyLabelText    = LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_KEY);
    const std::wstring colorLabelText  = LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_COLOR);

    const std::wstring pickButtonText       = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_PICK);
    const std::wstring setButtonText        = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_SET);
    const std::wstring clearButtonText      = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_CLEAR);
    const std::wstring loadFromFileText     = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_LOAD_FROM_FILE);
    const std::wstring duplicateThemeText   = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_DUPLICATE);
    const std::wstring saveThemeText        = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_SAVE_THEME);
    const std::wstring applyTemporarilyText = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_APPLY_TEMPORARILY);

    const DWORD baseStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    const DWORD wrapStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;
    const DWORD listExStyle     = state.theme.systemHighContrast ? WS_EX_CLIENTEDGE : 0;

    state.themesThemeLabel.reset(
        CreateWindowExW(0, L"Static", themeLabelText.c_str(), baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedComboBox(state, parent, state.themesThemeFrame, state.themesThemeCombo, IDC_PREFS_THEMES_THEME_COMBO);

    state.themesNameLabel.reset(
        CreateWindowExW(0, L"Static", nameLabelText.c_str(), baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedEditBox(
        state, parent, state.themesNameFrame, state.themesNameEdit, IDC_PREFS_THEMES_NAME_EDIT, WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);
    if (state.themesNameEdit)
    {
        SendMessageW(state.themesNameEdit.get(), EM_SETLIMITTEXT, 64, 0);
    }

    state.themesBaseLabel.reset(
        CreateWindowExW(0, L"Static", baseLabelText.c_str(), baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedComboBox(state, parent, state.themesBaseFrame, state.themesBaseCombo, IDC_PREFS_THEMES_BASE_COMBO);

    state.themesSearchLabel.reset(
        CreateWindowExW(0, L"Static", searchLabelText.c_str(), baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedEditBox(
        state, parent, state.themesSearchFrame, state.themesSearchEdit, IDC_PREFS_THEMES_SEARCH_EDIT, WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);
    if (state.themesSearchEdit)
    {
        SendMessageW(state.themesSearchEdit.get(), EM_SETLIMITTEXT, 128, 0);
    }

    state.themesColorsList.reset(CreateWindowExW(listExStyle,
                                                 WC_LISTVIEWW,
                                                 L"",
                                                 WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS | LVS_OWNERDRAWFIXED,
                                                 0,
                                                 0,
                                                 10,
                                                 10,
                                                 parent,
                                                 reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_THEMES_COLORS_LIST)),
                                                 GetModuleHandleW(nullptr),
                                                 nullptr));

    state.themesKeyLabel.reset(
        CreateWindowExW(0, L"Static", keyLabelText.c_str(), baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedEditBox(
        state, parent, state.themesKeyFrame, state.themesKeyEdit, IDC_PREFS_THEMES_KEY_EDIT, WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);
    if (state.themesKeyEdit)
    {
        SendMessageW(state.themesKeyEdit.get(), EM_SETLIMITTEXT, 64, 0);
    }

    state.themesColorLabel.reset(
        CreateWindowExW(0, L"Static", colorLabelText.c_str(), baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.themesColorSwatch.reset(CreateWindowExW(0,
                                                  L"Static",
                                                  L"",
                                                  WS_CHILD | WS_VISIBLE | SS_OWNERDRAW,
                                                  0,
                                                  0,
                                                  10,
                                                  10,
                                                  parent,
                                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_THEMES_COLOR_SWATCH)),
                                                  GetModuleHandleW(nullptr),
                                                  nullptr));
    PrefsInput::CreateFramedEditBox(
        state, parent, state.themesColorFrame, state.themesColorEdit, IDC_PREFS_THEMES_COLOR_EDIT, WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);
    if (state.themesColorEdit)
    {
        SendMessageW(state.themesColorEdit.get(), EM_SETLIMITTEXT, 11, 0); // "#AARRGGBB"
    }

    const bool customButtons     = ! state.theme.systemHighContrast;
    const DWORD themeButtonStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | (customButtons ? BS_OWNERDRAW : 0U);

    state.themesPickColor.reset(CreateWindowExW(0,
                                                L"Button",
                                                pickButtonText.c_str(),
                                                themeButtonStyle,
                                                0,
                                                0,
                                                10,
                                                10,
                                                parent,
                                                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_THEMES_PICK_COLOR)),
                                                GetModuleHandleW(nullptr),
                                                nullptr));
    state.themesSetOverride.reset(CreateWindowExW(0,
                                                  L"Button",
                                                  setButtonText.c_str(),
                                                  themeButtonStyle,
                                                  0,
                                                  0,
                                                  10,
                                                  10,
                                                  parent,
                                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_THEMES_SET_OVERRIDE)),
                                                  GetModuleHandleW(nullptr),
                                                  nullptr));
    state.themesRemoveOverride.reset(CreateWindowExW(0,
                                                     L"Button",
                                                     clearButtonText.c_str(),
                                                     themeButtonStyle,
                                                     0,
                                                     0,
                                                     10,
                                                     10,
                                                     parent,
                                                     reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_THEMES_REMOVE_OVERRIDE)),
                                                     GetModuleHandleW(nullptr),
                                                     nullptr));
    state.themesLoadFromFile.reset(CreateWindowExW(0,
                                                   L"Button",
                                                   loadFromFileText.c_str(),
                                                   themeButtonStyle,
                                                   0,
                                                   0,
                                                   10,
                                                   10,
                                                   parent,
                                                   reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_THEMES_LOAD_FILE)),
                                                   GetModuleHandleW(nullptr),
                                                   nullptr));
    state.themesDuplicateTheme.reset(CreateWindowExW(0,
                                                     L"Button",
                                                     duplicateThemeText.c_str(),
                                                     themeButtonStyle,
                                                     0,
                                                     0,
                                                     10,
                                                     10,
                                                     parent,
                                                     reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_THEMES_DUPLICATE_THEME)),
                                                     GetModuleHandleW(nullptr),
                                                     nullptr));
    state.themesSaveTheme.reset(CreateWindowExW(0,
                                                L"Button",
                                                saveThemeText.c_str(),
                                                themeButtonStyle,
                                                0,
                                                0,
                                                10,
                                                10,
                                                parent,
                                                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_THEMES_SAVE_THEME)),
                                                GetModuleHandleW(nullptr),
                                                nullptr));
    state.themesApplyTemporarily.reset(CreateWindowExW(0,
                                                       L"Button",
                                                       applyTemporarilyText.c_str(),
                                                       themeButtonStyle,
                                                       0,
                                                       0,
                                                       10,
                                                       10,
                                                       parent,
                                                       reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_THEMES_APPLY_TEMP)),
                                                       GetModuleHandleW(nullptr),
                                                       nullptr));

    state.themesNote.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
}

void ThemesPane::LayoutControls(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, int sectionY, HFONT dialogFont) noexcept
{
    if (! host)
    {
        return;
    }

    using namespace PrefsLayoutConstants;

    const UINT dpi = GetDpiForWindow(host);

    const int rowHeight   = std::max(1, ThemedControls::ScaleDip(dpi, kRowHeightDip));
    const int labelHeight = std::max(1, ThemedControls::ScaleDip(dpi, kTitleHeightDip));
    const int gapX        = ThemedControls::ScaleDip(dpi, kToggleGapXDip);

    const int themeLabelWidth = std::min(width, ThemedControls::ScaleDip(dpi, 60));
    const int editWidth       = std::max(0, width - themeLabelWidth - gapX);

    auto placeLabeledControl = [&](HWND label, HWND frame, HWND control, int controlWidth) noexcept
    {
        controlWidth           = std::max(0, std::min(editWidth, controlWidth));
        const int controlX     = x + themeLabelWidth + gapX;
        const int framePadding = (frame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;

        if (label)
        {
            SetWindowPos(label, nullptr, x, y + (rowHeight - labelHeight) / 2, themeLabelWidth, labelHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }
        if (frame)
        {
            SetWindowPos(frame, nullptr, controlX, y, controlWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        if (control)
        {
            SetWindowPos(control,
                         nullptr,
                         controlX + framePadding,
                         y + framePadding,
                         std::max(1, controlWidth - 2 * framePadding),
                         std::max(1, rowHeight - 2 * framePadding),
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }
        y += rowHeight + gapY;
    };

    int themeWidth          = state.themesThemeCombo ? ThemedControls::MeasureComboBoxPreferredWidth(state.themesThemeCombo.get(), dpi) : editWidth;
    const int minThemeWidth = ThemedControls::ScaleDip(dpi, 160);
    const int maxThemeWidth = std::max(minThemeWidth, std::max(0, editWidth));
    themeWidth              = std::clamp(themeWidth, minThemeWidth, maxThemeWidth);
    themeWidth              = std::min(themeWidth, ThemedControls::ScaleDip(dpi, 320));
    placeLabeledControl(state.themesThemeLabel.get(), state.themesThemeFrame.get(), state.themesThemeCombo.get(), themeWidth);
    if (state.themesThemeCombo)
    {
        ThemedControls::EnsureComboBoxDroppedWidth(state.themesThemeCombo.get(), dpi);
    }

    placeLabeledControl(state.themesNameLabel.get(), state.themesNameFrame.get(), state.themesNameEdit.get(), editWidth);

    int baseWidth = state.themesBaseCombo ? ThemedControls::MeasureComboBoxPreferredWidth(state.themesBaseCombo.get(), dpi) : editWidth;
    baseWidth     = std::max(baseWidth, ThemedControls::ScaleDip(dpi, 100));
    placeLabeledControl(state.themesBaseLabel.get(), state.themesBaseFrame.get(), state.themesBaseCombo.get(), baseWidth);
    if (state.themesBaseCombo)
    {
        ThemedControls::EnsureComboBoxDroppedWidth(state.themesBaseCombo.get(), dpi);
    }

    const int buttonHeight   = rowHeight;
    const int loadWidth      = std::min(width, ThemedControls::ScaleDip(dpi, 140));
    const int duplicateWidth = std::min(width, ThemedControls::ScaleDip(dpi, 110));
    const int saveWidth      = std::min(width, ThemedControls::ScaleDip(dpi, 120));
    const int applyWidth     = std::min(width, ThemedControls::ScaleDip(dpi, 150));

    int leftGroupWidth      = 0;
    auto addLeftButtonWidth = [&](HWND button, int buttonWidth) noexcept
    {
        if (! button)
        {
            return;
        }
        if (leftGroupWidth > 0)
        {
            leftGroupWidth += gapX;
        }
        leftGroupWidth += buttonWidth;
    };

    addLeftButtonWidth(state.themesLoadFromFile.get(), loadWidth);
    addLeftButtonWidth(state.themesDuplicateTheme.get(), duplicateWidth);
    addLeftButtonWidth(state.themesSaveTheme.get(), saveWidth);

    const bool wrapApply = state.themesApplyTemporarily && leftGroupWidth > 0 && (leftGroupWidth + gapX + applyWidth > width);
    const int row1Y      = y;
    const int row2Y      = row1Y + buttonHeight + gapY;

    int leftButtonsX = x;
    if (state.themesLoadFromFile)
    {
        SetWindowPos(state.themesLoadFromFile.get(), nullptr, leftButtonsX, row1Y, loadWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesLoadFromFile.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        leftButtonsX += loadWidth + gapX;
    }
    if (state.themesDuplicateTheme)
    {
        SetWindowPos(state.themesDuplicateTheme.get(), nullptr, leftButtonsX, row1Y, duplicateWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesDuplicateTheme.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        leftButtonsX += duplicateWidth + gapX;
    }
    if (state.themesSaveTheme)
    {
        SetWindowPos(state.themesSaveTheme.get(), nullptr, leftButtonsX, row1Y, saveWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesSaveTheme.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    if (state.themesApplyTemporarily)
    {
        const int applyX = x + width - applyWidth;
        const int applyY = wrapApply ? row2Y : row1Y;
        SetWindowPos(state.themesApplyTemporarily.get(), nullptr, applyX, applyY, applyWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesApplyTemporarily.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    y = wrapApply ? (row2Y + buttonHeight + gapY) : (row1Y + buttonHeight + gapY);

    if (state.themesNote)
    {
        const HFONT infoFont        = state.italicFont ? state.italicFont.get() : dialogFont;
        const std::wstring noteText = PrefsUi::GetWindowTextString(state.themesNote.get());
        const int noteHeight        = noteText.empty() ? 0 : PrefsUi::MeasureStaticTextHeight(host, infoFont, width, noteText);
        SetWindowPos(state.themesNote.get(), nullptr, x, y, width, std::max(0, noteHeight), SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesNote.get(), WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
        y += std::max(0, noteHeight) + sectionY;
    }

    const int searchLabelWidth   = std::min(width, ThemedControls::ScaleDip(dpi, 52));
    const int searchEditWidth    = std::max(0, width - searchLabelWidth - gapX);
    const int searchEditX        = x + searchLabelWidth + gapX;
    const int searchFramePadding = (state.themesSearchFrame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;

    if (state.themesSearchLabel)
    {
        SetWindowPos(
            state.themesSearchLabel.get(), nullptr, x, y + (rowHeight - labelHeight) / 2, searchLabelWidth, labelHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesSearchLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    if (state.themesSearchFrame)
    {
        SetWindowPos(state.themesSearchFrame.get(), nullptr, searchEditX, y, searchEditWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (state.themesSearchEdit)
    {
        SetWindowPos(state.themesSearchEdit.get(),
                     nullptr,
                     searchEditX + searchFramePadding,
                     y + searchFramePadding,
                     std::max(1, searchEditWidth - 2 * searchFramePadding),
                     std::max(1, rowHeight - 2 * searchFramePadding),
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesSearchEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    y += rowHeight + gapY;

    RECT hostClient{};
    GetClientRect(host, &hostClient);
    const int hostBottom        = std::max(0l, hostClient.bottom - hostClient.top);
    const int hostContentBottom = std::max(0, hostBottom - margin);

    const int editorHeight = rowHeight;
    const int editorTop    = std::max(y, hostContentBottom - editorHeight);
    const int listTop      = y;
    const int listBottom   = std::max(listTop, editorTop - gapY);
    const int listHeight   = std::max(0, listBottom - listTop);

    if (state.themesColorsList)
    {
        SetWindowPos(state.themesColorsList.get(), nullptr, x, listTop, width, listHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesColorsList.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        UpdateThemesColorsListColumnWidths(state.themesColorsList.get(), dpi);
    }

    const int keyLabelWidth   = std::min(width, ThemedControls::ScaleDip(dpi, 34));
    const int colorLabelWidth = std::min(width, ThemedControls::ScaleDip(dpi, 44));
    const int pickWidth       = std::min(width, ThemedControls::ScaleDip(dpi, 70));
    const int setWidth        = std::min(width, ThemedControls::ScaleDip(dpi, 60));
    const int clearWidth      = std::min(width, ThemedControls::ScaleDip(dpi, 70));
    const int swatchWidth     = std::min(width, ThemedControls::ScaleDip(dpi, 22));
    const int colorEditWidth  = std::min(width, ThemedControls::ScaleDip(dpi, 110));

    const int buttonsWidth = pickWidth + gapX + setWidth + gapX + clearWidth;
    const int editAreaWidth =
        std::max(0, width - keyLabelWidth - gapX - colorLabelWidth - gapX - swatchWidth - gapX - colorEditWidth - gapX - buttonsWidth - gapX);

    if (state.themesKeyLabel)
    {
        SetWindowPos(
            state.themesKeyLabel.get(), nullptr, x, editorTop + (rowHeight - labelHeight) / 2, keyLabelWidth, labelHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesKeyLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    const int keyEditX        = x + keyLabelWidth + gapX;
    const int keyFramePadding = (state.themesKeyFrame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;
    if (state.themesKeyFrame)
    {
        SetWindowPos(state.themesKeyFrame.get(), nullptr, keyEditX, editorTop, editAreaWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (state.themesKeyEdit)
    {
        SetWindowPos(state.themesKeyEdit.get(),
                     nullptr,
                     keyEditX + keyFramePadding,
                     editorTop + keyFramePadding,
                     std::max(1, editAreaWidth - 2 * keyFramePadding),
                     std::max(1, rowHeight - 2 * keyFramePadding),
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesKeyEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    const int colorLabelX = x + keyLabelWidth + gapX + editAreaWidth + gapX;
    if (state.themesColorLabel)
    {
        SetWindowPos(state.themesColorLabel.get(),
                     nullptr,
                     colorLabelX,
                     editorTop + (rowHeight - labelHeight) / 2,
                     colorLabelWidth,
                     labelHeight,
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesColorLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    const int colorSwatchX = colorLabelX + colorLabelWidth + gapX;
    if (state.themesColorSwatch)
    {
        SetWindowPos(state.themesColorSwatch.get(), nullptr, colorSwatchX, editorTop, swatchWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    const int colorEditX        = colorSwatchX + swatchWidth + gapX;
    const int colorFramePadding = (state.themesColorFrame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;
    if (state.themesColorFrame)
    {
        SetWindowPos(state.themesColorFrame.get(), nullptr, colorEditX, editorTop, colorEditWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (state.themesColorEdit)
    {
        SetWindowPos(state.themesColorEdit.get(),
                     nullptr,
                     colorEditX + colorFramePadding,
                     editorTop + colorFramePadding,
                     std::max(1, colorEditWidth - 2 * colorFramePadding),
                     std::max(1, rowHeight - 2 * colorFramePadding),
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesColorEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    int buttonX = colorEditX + colorEditWidth + gapX;
    if (state.themesPickColor)
    {
        SetWindowPos(state.themesPickColor.get(), nullptr, buttonX, editorTop, pickWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesPickColor.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        buttonX += pickWidth + gapX;
    }
    if (state.themesSetOverride)
    {
        SetWindowPos(state.themesSetOverride.get(), nullptr, buttonX, editorTop, setWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesSetOverride.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        buttonX += setWidth + gapX;
    }
    if (state.themesRemoveOverride)
    {
        SetWindowPos(state.themesRemoveOverride.get(), nullptr, buttonX, editorTop, clearWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.themesRemoveOverride.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
}

void ThemesPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    RefreshThemesPage(host, state);
}

bool ThemesPane::HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND /*hwndCtl*/) noexcept
{
    switch (commandId)
    {
        case IDC_PREFS_THEMES_SEARCH_EDIT:
            if (notifyCode == EN_CHANGE)
            {
                const auto themeIdOpt = TryGetSelectedThemeId(state);
                if (! themeIdOpt.has_value())
                {
                    return true;
                }

                bool editable   = false;
                const auto* def = FindThemeDefinitionForDisplay(state, themeIdOpt.value(), editable);
                RefreshThemesColorsList(host, state, themeIdOpt.value(), def);
                ThemesPane::UpdateEditorFromSelection(host, state);
                if (state.themesColorsList)
                {
                    InvalidateRect(state.themesColorsList.get(), nullptr, FALSE);
                }
                return true;
            }
            break;

        case IDC_PREFS_THEMES_THEME_COMBO:
            if (notifyCode == CBN_SELCHANGE)
            {
                if (state.refreshingThemesPage)
                {
                    return true;
                }

                const ThemeComboItem* selected = TryGetSelectedThemeComboItem(state);
                if (! selected)
                {
                    return true;
                }

                if (selected->source == ThemeSchemaSource::New)
                {
                    BeginNewThemeCreation(host, state);
                    return true;
                }

                state.workingSettings.theme.currentThemeId.assign(selected->id);
                SetDirty(GetParent(host), state);
                RefreshThemesPage(host, state);
                return true;
            }
            break;

        case IDC_PREFS_THEMES_BASE_COMBO:
            if (notifyCode == CBN_SELCHANGE)
            {
                if (state.refreshingThemesPage)
                {
                    return true;
                }

                const auto themeIdOpt = TryGetSelectedThemeId(state);
                if (! themeIdOpt.has_value())
                {
                    return true;
                }

                auto* def = FindWorkingThemeDefinition(state, themeIdOpt.value());
                if (! def)
                {
                    return true;
                }

                const LRESULT sel = SendMessageW(state.themesBaseCombo.get(), CB_GETCURSEL, 0, 0);
                if (sel == CB_ERR)
                {
                    return true;
                }

                const LRESULT data = SendMessageW(state.themesBaseCombo.get(), CB_GETITEMDATA, static_cast<WPARAM>(sel), 0);
                if (data == CB_ERR)
                {
                    return true;
                }

                if (data < 0)
                {
                    def->baseThemeId.assign(L"builtin/system");
                }
                else
                {
                    const size_t optionIndex = static_cast<size_t>(data);
                    if (optionIndex >= kBuiltinThemeOptions.size())
                    {
                        return true;
                    }

                    def->baseThemeId.assign(kBuiltinThemeOptions[optionIndex].id);
                }

                SetDirty(GetParent(host), state);
                RefreshThemesPage(host, state);
                return true;
            }
            break;

        case IDC_PREFS_THEMES_NAME_EDIT:
            if (notifyCode == EN_CHANGE)
            {
                if (state.refreshingThemesPage)
                {
                    return true;
                }

                const auto themeIdOpt = TryGetSelectedThemeId(state);
                if (! themeIdOpt.has_value())
                {
                    return true;
                }

                auto* def = FindWorkingThemeDefinition(state, themeIdOpt.value());
                if (! def)
                {
                    return true;
                }

                const std::wstring name = PrefsUi::GetWindowTextString(state.themesNameEdit.get());
                if (name.empty())
                {
                    return true;
                }

                def->name = name;

                SetDirty(GetParent(host), state);
                return true;
            }
            if (notifyCode == EN_KILLFOCUS)
            {
                if (state.refreshingThemesPage)
                {
                    return true;
                }

                SyncSelectedUserThemeIdToName(host, state);
                return true;
            }
            break;

        case IDC_PREFS_THEMES_COLOR_EDIT:
            if (notifyCode == EN_CHANGE)
            {
                if (state.themesColorSwatch)
                {
                    InvalidateRect(state.themesColorSwatch.get(), nullptr, TRUE);
                }
                return true;
            }
            break;

        case IDC_PREFS_THEMES_PICK_COLOR:
            if (notifyCode == BN_CLICKED)
            {
                PickThemeColorIntoEditor(host, state);
                return true;
            }
            break;

        case IDC_PREFS_THEMES_SET_OVERRIDE:
            if (notifyCode == BN_CLICKED)
            {
                SetThemeOverrideFromEditor(host, state);
                return true;
            }
            break;

        case IDC_PREFS_THEMES_REMOVE_OVERRIDE:
            if (notifyCode == BN_CLICKED)
            {
                ClearThemeOverrideFromEditor(host, state);
                return true;
            }
            break;

        case IDC_PREFS_THEMES_LOAD_FILE:
            if (notifyCode == BN_CLICKED)
            {
                LoadThemeFromFile(host, state);
                return true;
            }
            break;

        case IDC_PREFS_THEMES_DUPLICATE_THEME:
            if (notifyCode == BN_CLICKED)
            {
                DuplicateSelectedTheme(host, state);
                return true;
            }
            break;

        case IDC_PREFS_THEMES_SAVE_THEME:
            if (notifyCode == BN_CLICKED)
            {
                SaveThemeToFile(host, state);
                return true;
            }
            break;

        case IDC_PREFS_THEMES_APPLY_TEMP:
            if (notifyCode == BN_CLICKED)
            {
                ApplyThemeTemporarily(host, state);
                return true;
            }
            break;
    }

    return false;
}

void ThemesPane::UpdateEditorFromSelection(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    if (! state.themesKeyEdit || ! state.themesColorEdit)
    {
        return;
    }

    if (! state.themesColorsList)
    {
        return;
    }

    const int selected = ListView_GetNextItem(state.themesColorsList.get(), -1, LVNI_SELECTED);
    if (selected < 0)
    {
        state.refreshingThemesPage = true;
        const auto reset           = wil::scope_exit([&] { state.refreshingThemesPage = false; });
        SetWindowTextW(state.themesKeyEdit.get(), L"");
        SetWindowTextW(state.themesColorEdit.get(), L"");
        if (state.themesColorSwatch)
        {
            InvalidateRect(state.themesColorSwatch.get(), nullptr, TRUE);
        }
        return;
    }

    wchar_t keyText[128]{};
    ListView_GetItemText(state.themesColorsList.get(), selected, 0, keyText, static_cast<int>(std::size(keyText)));

    wchar_t valueText[64]{};
    ListView_GetItemText(state.themesColorsList.get(), selected, 1, valueText, static_cast<int>(std::size(valueText)));

    state.refreshingThemesPage = true;
    const auto reset           = wil::scope_exit([&] { state.refreshingThemesPage = false; });
    SetWindowTextW(state.themesKeyEdit.get(), keyText);
    SetWindowTextW(state.themesColorEdit.get(), valueText);
    if (state.themesColorSwatch)
    {
        InvalidateRect(state.themesColorSwatch.get(), nullptr, TRUE);
    }
}

bool ThemesPane::HandleNotify(HWND host, PreferencesDialogState& state, NMHDR* hdr, LRESULT& outResult) noexcept
{
    if (! hdr || ! state.themesColorsList || hdr->hwndFrom != state.themesColorsList.get())
    {
        return false;
    }

    switch (hdr->code)
    {
        case NM_CUSTOMDRAW: outResult = CDRF_DODEFAULT; return true;
        case NM_SETFOCUS:
            PrefsPaneHost::EnsureControlVisible(host, state, state.themesColorsList.get());
            InvalidateRect(state.themesColorsList.get(), nullptr, FALSE);
            outResult = 0;
            return true;
        case NM_KILLFOCUS:
            InvalidateRect(state.themesColorsList.get(), nullptr, FALSE);
            outResult = 0;
            return true;
        case LVN_ITEMCHANGED:
            ThemesPane::UpdateEditorFromSelection(host, state);
            outResult = 0;
            return true;
    }

    return false;
}

LRESULT ThemesPane::OnMeasureColorsList(MEASUREITEMSTRUCT* mis, PreferencesDialogState& state) noexcept
{
    if (! mis || mis->CtlType != ODT_LISTVIEW || mis->CtlID != static_cast<UINT>(IDC_PREFS_THEMES_COLORS_LIST))
    {
        return 0;
    }

    if (! state.themesColorsList)
    {
        return 0;
    }

    wil::unique_hdc_window hdc(GetDC(state.themesColorsList.get()));
    if (! hdc)
    {
        mis->itemHeight = 26u;
        return 1;
    }

    const HFONT font = reinterpret_cast<HFONT>(SendMessageW(state.themesColorsList.get(), WM_GETFONT, 0, 0));
    if (font)
    {
        [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), font);
        mis->itemHeight               = static_cast<UINT>(std::max(1, PrefsListView::GetSingleLineRowHeightPx(state.themesColorsList.get(), hdc.get())));
        return 1;
    }

    mis->itemHeight = 26u;
    return 1;
}

LRESULT ThemesPane::OnDrawColorsList(DRAWITEMSTRUCT* dis, PreferencesDialogState& state) noexcept
{
    if (! dis || dis->CtlType != ODT_LISTVIEW || dis->CtlID != static_cast<UINT>(IDC_PREFS_THEMES_COLORS_LIST))
    {
        return 0;
    }

    if (! state.themesColorsList || ! dis->hDC)
    {
        return 1;
    }

    const int itemIndex = static_cast<int>(dis->itemID);
    if (itemIndex < 0)
    {
        return 1;
    }

    RECT rc = dis->rcItem;
    if (rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return 1;
    }

    wchar_t seedText[256]{};
    ListView_GetItemText(state.themesColorsList.get(), itemIndex, 0, seedText, static_cast<int>(std::size(seedText)));
    const std::wstring_view seed = std::wstring_view(seedText, std::wcslen(seedText));

    const bool selected    = (dis->itemState & ODS_SELECTED) != 0;
    const bool focused     = (dis->itemState & ODS_FOCUS) != 0;
    const bool listFocused = GetFocus() == state.themesColorsList.get();

    const HWND root         = GetAncestor(state.themesColorsList.get(), GA_ROOT);
    const bool windowActive = root && GetActiveWindow() == root;

    COLORREF bg        = state.theme.systemHighContrast ? GetSysColor(COLOR_WINDOW) : state.theme.windowBackground;
    COLORREF textColor = state.theme.systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT) : state.theme.menu.text;

    if (selected)
    {
        COLORREF selBg = state.theme.systemHighContrast ? GetSysColor(COLOR_HIGHLIGHT) : state.theme.menu.selectionBg;
        if (! state.theme.highContrast && state.theme.menu.rainbowMode && ! seed.empty())
        {
            selBg = RainbowMenuSelectionColor(seed, state.theme.menu.darkBase);
        }

        COLORREF selText = state.theme.systemHighContrast ? GetSysColor(COLOR_HIGHLIGHTTEXT) : state.theme.menu.selectionText;
        if (! state.theme.highContrast && state.theme.menu.rainbowMode)
        {
            selText = ChooseContrastingTextColor(selBg);
        }

        if (windowActive && listFocused)
        {
            bg        = selBg;
            textColor = selText;
        }
        else if (! state.theme.highContrast)
        {
            const int denom = state.theme.menu.darkBase ? 2 : 3;
            bg              = ThemedControls::BlendColor(state.theme.windowBackground, selBg, 1, denom);
            textColor       = ChooseContrastingTextColor(bg);
        }
        else
        {
            bg        = selBg;
            textColor = selText;
        }
    }
    else if (! state.theme.highContrast && ((itemIndex % 2) == 1))
    {
        const COLORREF tint =
            (state.theme.menu.rainbowMode && ! seed.empty()) ? RainbowMenuSelectionColor(seed, state.theme.menu.darkBase) : state.theme.menu.selectionBg;
        const int denom = state.theme.menu.darkBase ? 6 : 8;
        bg              = ThemedControls::BlendColor(bg, tint, 1, denom);
    }

    wil::unique_hbrush bgBrush(CreateSolidBrush(bg));
    if (bgBrush)
    {
        FillRect(dis->hDC, &rc, bgBrush.get());
    }

    if (! state.theme.highContrast && textColor == bg)
    {
        textColor = ChooseContrastingTextColor(bg);
    }

    const UINT dpi     = GetDpiForWindow(state.themesColorsList.get());
    const int paddingX = ThemedControls::ScaleDip(dpi, 8);

    const int col0W = std::max(0, ListView_GetColumnWidth(state.themesColorsList.get(), 0));
    const int col1W = std::max(0, ListView_GetColumnWidth(state.themesColorsList.get(), 1));
    const int col2W = std::max(0, ListView_GetColumnWidth(state.themesColorsList.get(), 2));

    RECT col0Rect  = rc;
    col0Rect.right = std::min(rc.right, rc.left + col0W);

    RECT col1Rect  = rc;
    col1Rect.left  = col0Rect.right;
    col1Rect.right = (col1W > 0) ? std::min(rc.right, col1Rect.left + col1W) : rc.right;

    RECT col2Rect  = rc;
    col2Rect.left  = col1Rect.right;
    col2Rect.right = (col2W > 0) ? std::min(rc.right, col2Rect.left + col2W) : rc.right;

    wchar_t text0[256]{};
    ListView_GetItemText(state.themesColorsList.get(), itemIndex, 0, text0, static_cast<int>(std::size(text0)));
    wchar_t text1[512]{};
    ListView_GetItemText(state.themesColorsList.get(), itemIndex, 1, text1, static_cast<int>(std::size(text1)));

    bool overridden = false;
    LVITEMW paramItem{};
    paramItem.mask     = LVIF_PARAM;
    paramItem.iItem    = itemIndex;
    paramItem.iSubItem = 0;
    if (ListView_GetItem(state.themesColorsList.get(), &paramItem))
    {
        overridden = paramItem.lParam != 0;
    }

    HFONT normalFont = reinterpret_cast<HFONT>(SendMessageW(state.themesColorsList.get(), WM_GETFONT, 0, 0));
    if (! normalFont)
    {
        normalFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    const HFONT boldFont          = (overridden && state.boldFont) ? state.boldFont.get() : normalFont;
    [[maybe_unused]] auto oldFont = wil::SelectObject(dis->hDC, normalFont);

    SetBkMode(dis->hDC, TRANSPARENT);
    SetTextColor(dis->hDC, textColor);

    RECT textRect0  = col0Rect;
    textRect0.left  = std::min(textRect0.right, textRect0.left + paddingX);
    textRect0.right = std::max(textRect0.left, textRect0.right - paddingX);

    if (overridden && boldFont && boldFont != normalFont)
    {
        [[maybe_unused]] auto oldKeyFont = wil::SelectObject(dis->hDC, boldFont);
        DrawTextW(dis->hDC, text0, static_cast<int>(std::wcslen(text0)), &textRect0, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
    }
    else
    {
        DrawTextW(dis->hDC, text0, static_cast<int>(std::wcslen(text0)), &textRect0, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
    }

    RECT textRect1  = col1Rect;
    textRect1.left  = std::min(textRect1.right, textRect1.left + paddingX);
    textRect1.right = std::max(textRect1.left, textRect1.right - paddingX);

    DrawTextW(dis->hDC, text1, static_cast<int>(std::wcslen(text1)), &textRect1, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);

    std::optional<uint32_t> argb;
    const std::wstring_view valueText = std::wstring_view(text1, std::wcslen(text1));
    uint32_t parsed                   = 0;
    if (! valueText.empty() && Common::Settings::TryParseColor(valueText, parsed))
    {
        argb = parsed;
    }

    RECT swatchRect         = col2Rect;
    const int swatchPadding = ThemedControls::ScaleDip(dpi, 4);
    InflateRect(&swatchRect, -swatchPadding, -swatchPadding);
    const int swatchW    = std::max(0l, swatchRect.right - swatchRect.left);
    const int swatchH    = std::max(0l, swatchRect.bottom - swatchRect.top);
    const int swatchSize = std::min(swatchW, swatchH);
    if (swatchSize > 0)
    {
        swatchRect.left += (swatchW - swatchSize) / 2;
        swatchRect.top += (swatchH - swatchSize) / 2;
        swatchRect.right  = swatchRect.left + swatchSize;
        swatchRect.bottom = swatchRect.top + swatchSize;
        DrawRoundedColorSwatch(dis->hDC, swatchRect, dpi, state.theme, bg, argb, true);
    }

    if (focused)
    {
        RECT focusRc = rc;
        InflateRect(&focusRc,
                    -ThemedControls::ScaleDip(dpi, PrefsLayoutConstants::kFramePaddingDip),
                    -ThemedControls::ScaleDip(dpi, PrefsLayoutConstants::kFramePaddingDip));

        COLORREF focusTint = state.theme.menu.selectionBg;
        if (! state.theme.highContrast && state.theme.menu.rainbowMode && ! seed.empty())
        {
            focusTint = RainbowMenuSelectionColor(seed, state.theme.menu.darkBase);
        }

        const int weight          = (windowActive && listFocused) ? (state.theme.dark ? 70 : 55) : (state.theme.dark ? 55 : 40);
        const COLORREF focusColor = state.theme.systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT) : ThemedControls::BlendColor(bg, focusTint, weight, 255);

        wil::unique_hpen focusPen(CreatePen(PS_SOLID, 1, focusColor));
        if (focusPen)
        {
            [[maybe_unused]] auto oldBrush2 = wil::SelectObject(dis->hDC, GetStockObject(NULL_BRUSH));
            [[maybe_unused]] auto oldPen2   = wil::SelectObject(dis->hDC, focusPen.get());
            Rectangle(dis->hDC, focusRc.left, focusRc.top, focusRc.right, focusRc.bottom);
        }
    }

    return 1;
}

LRESULT ThemesPane::OnDrawColorSwatch(DRAWITEMSTRUCT* dis, PreferencesDialogState& state) noexcept
{
    if (! dis || dis->CtlType != ODT_STATIC || dis->CtlID != static_cast<UINT>(IDC_PREFS_THEMES_COLOR_SWATCH))
    {
        return 0;
    }

    if (! dis->hwndItem || ! dis->hDC)
    {
        return 1;
    }

    const UINT dpi    = GetDpiForWindow(dis->hwndItem);
    const COLORREF bg = state.theme.systemHighContrast ? GetSysColor(COLOR_WINDOW) : state.theme.windowBackground;

    wil::unique_hbrush bgBrushOwned;
    HBRUSH bgBrush = state.backgroundBrush ? state.backgroundBrush.get() : nullptr;
    if (! bgBrush)
    {
        bgBrushOwned.reset(CreateSolidBrush(bg));
        bgBrush = bgBrushOwned.get();
    }
    if (bgBrush)
    {
        FillRect(dis->hDC, &dis->rcItem, bgBrush);
    }

    std::optional<uint32_t> argb;
    if (state.themesColorEdit)
    {
        const std::wstring valueText = PrefsUi::GetWindowTextString(state.themesColorEdit.get());
        uint32_t parsed              = 0;
        if (! valueText.empty() && Common::Settings::TryParseColor(valueText, parsed))
        {
            argb = parsed;
        }
    }

    RECT swatch = dis->rcItem;
    InflateRect(&swatch,
                -ThemedControls::ScaleDip(dpi, PrefsLayoutConstants::kFramePaddingDip),
                -ThemedControls::ScaleDip(dpi, PrefsLayoutConstants::kFramePaddingDip));
    DrawRoundedColorSwatch(dis->hDC, swatch, dpi, state.theme, bg, argb, IsWindowEnabled(dis->hwndItem) != FALSE);
    return 1;
}
