// Preferences.Themes.cpp

#include "Framework.h"

#include "Preferences.Themes.h"

#include <algorithm>
#include <array>
#include <cstdlib>
#include <cwchar>
#include <cwctype>
#include <filesystem>
#include <format>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include <commdlg.h>
#include <uxtheme.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

#include "D2DHdcPaint.h"
#include "Helpers.h"
#include "HostServices.h"
#include "UiMetrics.h"
#include "WindowMessages.h"
#include "resource.h"

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ColorSwatch;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::Control;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridCellKind;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::WindowHost;

#ifdef ENABLE_TESTS
enum class DebugThemeBrowseResultKind
{
    Path,
    Cancel,
};

struct DebugThemeBrowseResult
{
    DebugThemeBrowseResultKind kind = DebugThemeBrowseResultKind::Path;
    std::filesystem::path path{};
};

std::mutex g_debugThemeBrowseResultMutex;
std::optional<DebugThemeBrowseResult> g_debugNextThemeBrowseResult;
#endif

void LogThemesDxState(const wchar_t* reason,
                      HWND pageHostWindow,
                      HWND hostWindow,
                      const WindowHost* host,
                      const Panel* pageContentRoot,
                      const void* dxState,
                      bool rebuildOnShow) noexcept
{
    size_t wrapperChildren = 0u;
    if (pageContentRoot)
    {
        wrapperChildren = pageContentRoot->GetChildren().size();
    }

    Debug::Info(L"Preferences.Themes: reason={} pageHostWindow={:#x} hostWindow={:#x} dxHost={} root={} dxState={} wrapperChildren={} focus={} bridge={} "
                L"dx={}x{} renderCount={} resizeCount={} resizeFailures={} rebuildOnShow={}",
                reason ? reason : L"(null)",
                reinterpret_cast<uintptr_t>(pageHostWindow),
                reinterpret_cast<uintptr_t>(hostWindow),
                static_cast<const void*>(host),
                static_cast<const void*>(pageContentRoot),
                dxState,
                wrapperChildren,
                host ? static_cast<const void*>(host->GetFocusControl()) : nullptr,
                (host && host->HasActiveTextInputBridge()) ? L"true" : L"false",
                GetDxHostDebugWidthPx(host),
                GetDxHostDebugHeightPx(host),
                GetDxHostDebugRenderCount(host),
                GetDxHostDebugResizeCount(host),
                GetDxHostDebugResizeFailureCount(host),
                rebuildOnShow ? L"true" : L"false");
}

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
    return UiMetrics::BlendColor(background, rgb, alpha, 255);
}

void DrawRoundedColorSwatch(HDC hdc, RECT rc, UINT dpi, const AppTheme& theme, COLORREF background, std::optional<uint32_t> argb, bool enabled) noexcept
{
    if (! hdc || rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return;
    }

    const int width  = std::max(0l, rc.right - rc.left);
    const int height = std::max(0l, rc.bottom - rc.top);
    const int radius = std::max(1, std::min(UiMetrics::ScaleDip(dpi, 4), std::min(width, height) / 2));

    COLORREF border = theme.systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT) : UiMetrics::BlendColor(background, theme.menu.text, theme.dark ? 70 : 50, 255);
    COLORREF fill   = background;
    if (argb.has_value())
    {
        fill = CompositeArgbOnBackground(background, argb.value());
    }

    if (! enabled && ! theme.highContrast)
    {
        fill   = UiMetrics::BlendColor(background, fill, theme.dark ? 120 : 95, 255);
        border = UiMetrics::BlendColor(background, border, theme.dark ? 120 : 95, 255);
    }

    D2DHdcPaint::Session paint;
    if (! paint.Begin(hdc, rc))
    {
        return;
    }
    paint.FillRoundedRectangle(rc, static_cast<float>(radius), fill, border);
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

[[nodiscard]] uint64_t MakeThemesStableRowId(std::wstring_view key) noexcept
{
    constexpr uint64_t kFNVOffset = 1469598103934665603ull;
    constexpr uint64_t kFNVPrime  = 1099511628211ull;

    uint64_t value = kFNVOffset;
    for (const wchar_t ch : key)
    {
        value ^= static_cast<uint64_t>(std::towlower(static_cast<wint_t>(ch)));
        value *= kFNVPrime;
    }
    return value;
}

void SetThemesNoteText(PreferencesDialogState& state, std::wstring text) noexcept
{
    state.themesNoteText = std::move(text);
}

void SetThemesNameText(PreferencesDialogState& state, std::wstring text) noexcept
{
    state.themesNameText = std::move(text);
}

void SetThemesKeyText(PreferencesDialogState& state, std::wstring text) noexcept
{
    state.themesKeyText = std::move(text);
}

void SetThemesColorText(PreferencesDialogState& state, std::wstring text) noexcept
{
    state.themesColorText = std::move(text);
}

struct ThemesGridRow
{
    uint64_t stableId = 0u;
    std::wstring key;
    std::wstring value;
    bool overridden    = false;
    bool hasColor      = false;
    uint32_t colorArgb = 0u;
};

class ThemesGridModel final : public IDxGridModel
{
public:
    ThemesGridModel()
    {
        _columns = {
            {L"key", LoadStringResource(nullptr, IDS_PREFS_THEMES_COL_KEY), 260.0f, 120.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
            {L"value", LoadStringResource(nullptr, IDS_PREFS_THEMES_COL_VALUE), 140.0f, 96.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
            {L"swatch", L"", 44.0f, 36.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false, DWRITE_TEXT_ALIGNMENT_CENTER},
        };
    }

    void SetRows(std::vector<ThemesGridRow> rows)
    {
        _rows = std::move(rows);
        _rowIndexByStableId.clear();
        _rowIndexByStableId.reserve(_rows.size());
        for (size_t rowIndex = 0u; rowIndex < _rows.size(); ++rowIndex)
        {
            _rowIndexByStableId[_rows[rowIndex].stableId] = rowIndex;
        }
    }

    [[nodiscard]] const std::vector<ThemesGridRow>& GetRows() const noexcept
    {
        return _rows;
    }

    [[nodiscard]] std::optional<size_t> FindRowIndexByKey(std::wstring_view key) const noexcept
    {
        const auto it = std::find_if(_rows.begin(), _rows.end(), [&](const ThemesGridRow& row) noexcept { return row.key == key; });
        if (it == _rows.end())
        {
            return std::nullopt;
        }
        return static_cast<size_t>(std::distance(_rows.begin(), it));
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rows.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return _columns.at(columnIndex);
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        if (rowIndex >= _rows.size())
        {
            return;
        }

        const ThemesGridRow& row = _rows[rowIndex];
        switch (columnIndex)
        {
            case 0:
                outCell.text = row.key;
                if (row.overridden)
                {
                    outCell.badgeText = L"*";
                    outCell.badgeTone = RedSalamander::DxUi::AdornmentTone::Accent;
                }
                break;
            case 1: outCell.text = row.value; break;
            case 2:
                outCell.kind           = GridCellKind::ColorSwatch;
                outCell.hasSwatchValue = row.hasColor;
                outCell.swatchArgb     = row.colorArgb;
                outCell.textAlignment  = DWRITE_TEXT_ALIGNMENT_CENTER;
                break;
            default: break;
        }
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        if (rowIndex >= _rows.size())
        {
            return 0u;
        }
        return _rows[rowIndex].stableId;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        const auto it = _rowIndexByStableId.find(rowId);
        if (it == _rowIndexByStableId.end())
        {
            return std::nullopt;
        }

        return it->second;
    }

private:
    std::vector<GridColumnDesc> _columns;
    std::vector<ThemesGridRow> _rows;
    std::unordered_map<uint64_t, size_t> _rowIndexByStableId;
};

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

#ifdef ENABLE_TESTS
    {
        std::scoped_lock lock(g_debugThemeBrowseResultMutex);
        if (g_debugNextThemeBrowseResult.has_value())
        {
            const DebugThemeBrowseResult result = *g_debugNextThemeBrowseResult;
            g_debugNextThemeBrowseResult.reset();
            if (result.kind == DebugThemeBrowseResultKind::Cancel)
            {
                return false;
            }

            outPath = result.path;
            return ! outPath.empty();
        }
    }
#endif

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

#ifdef ENABLE_TESTS
bool DebugSetPreferencesThemesNextBrowsePathImpl(const std::wstring_view path) noexcept
{
    std::scoped_lock lock(g_debugThemeBrowseResultMutex);
    if (path.empty())
    {
        g_debugNextThemeBrowseResult.reset();
        return true;
    }

    g_debugNextThemeBrowseResult = DebugThemeBrowseResult{.kind = DebugThemeBrowseResultKind::Path, .path = std::filesystem::path(path)};
    return true;
}

bool DebugCancelPreferencesThemesNextBrowseImpl() noexcept
{
    std::scoped_lock lock(g_debugThemeBrowseResultMutex);
    g_debugNextThemeBrowseResult = DebugThemeBrowseResult{.kind = DebugThemeBrowseResultKind::Cancel};
    return true;
}
#endif

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
    EnsureThemeFileThemesLoaded(state);

    state.themeComboItems.clear();

    auto addTheme = [&](std::wstring_view id, std::wstring_view name, ThemeSchemaSource source) noexcept
    {
        ThemeComboItem item;
        item.id          = std::wstring(id);
        item.displayName = name.empty() ? std::wstring(id) : std::wstring(name);
        item.source      = source;

        state.themeComboItems.push_back(std::move(item));
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
}

[[nodiscard]] const ThemeComboItem* FindThemeComboItemById(const PreferencesDialogState& state, std::wstring_view id) noexcept
{
    if (id.empty())
    {
        return nullptr;
    }

    const auto it = std::find_if(
        state.themeComboItems.begin(), state.themeComboItems.end(), [&](const ThemeComboItem& item) noexcept { return std::wstring_view(item.id) == id; });
    return it != state.themeComboItems.end() ? &(*it) : nullptr;
}

[[nodiscard]] const ThemeComboItem* TryGetSelectedThemeComboItem(const PreferencesDialogState& state) noexcept
{
    if (const auto* retained = FindThemeComboItemById(state, state.workingSettings.theme.currentThemeId))
    {
        return retained;
    }

    return nullptr;
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

[[nodiscard]] const Common::Settings::ThemeDefinition* FindThemeDefinitionForDisplay(const PreferencesDialogState& state,
                                                                                     std::wstring_view id,
                                                                                     bool& outEditable) noexcept
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
    D2D1::ColorF metaPerf        = D2D1::ColorF(D2D1::ColorF::MediumSeaGreen);
    D2D1::ColorF metaDebug       = D2D1::ColorF(D2D1::ColorF::MediumPurple);
};
[[nodiscard]] MonitorTextViewTheme ResolveMonitorThemeForDisplay(std::wstring_view baseThemeId,
                                                                 const std::unordered_map<std::wstring, uint32_t>* overrides) noexcept;
[[nodiscard]] std::optional<uint32_t> TryGetEffectiveThemeColorArgb(const AppTheme& appTheme,
                                                                    const MonitorTextViewTheme& monitorTheme,
                                                                    const std::unordered_map<std::wstring, uint32_t>* overrides,
                                                                    std::wstring_view key) noexcept;

constexpr std::array<std::wstring_view, 65> kKnownColorKeys = {{
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
    L"monitor.textView.metaPerf",
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

    L"viewer.diff.addedBackground",
    L"viewer.diff.removedBackground",
    L"viewer.diff.contextBackground",
    L"viewer.diff.headerBackground",
    L"viewer.diff.bannerBackground",
    L"viewer.diff.placeholderBackground",
    L"viewer.diff.divider",
}};

[[nodiscard]] std::vector<ThemesGridRow> BuildThemesColorRows(const PreferencesDialogState& state) noexcept
{
    std::vector<ThemesGridRow> rows;

    const auto themeIdOpt = TryGetSelectedThemeId(state);
    if (! themeIdOpt.has_value())
    {
        return rows;
    }

    const std::wstring_view themeId     = themeIdOpt.value();
    bool editable                       = false;
    const auto* def                     = FindThemeDefinitionForDisplay(state, themeId, editable);
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

    const std::wstring_view filter = PrefsUi::TrimWhitespace(state.themesSearchText);

    std::vector<std::wstring> extraKeys;
    if (overrides)
    {
        extraKeys.reserve(overrides->size());
        for (const auto& [key, _] : *overrides)
        {
            bool known = false;
            for (const auto knownKey : kKnownColorKeys)
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
    allKeys.reserve(kKnownColorKeys.size() + extraKeys.size());
    for (const auto key : kKnownColorKeys)
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

        ThemesGridRow row{};
        row.key        = key;
        row.value      = Common::Settings::FormatColor(valueOpt.value());
        row.stableId   = MakeThemesStableRowId(row.key);
        row.overridden = overrides != nullptr && overrides->contains(key);
        row.hasColor   = true;
        row.colorArgb  = valueOpt.value();
        rows.push_back(std::move(row));
    }

    return rows;
}

void RefreshThemesPage(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

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

    if (editable)
    {
        SetThemesNoteText(state, L"");
    }
    else if (def)
    {
        SetThemesNoteText(state, LoadStringResource(nullptr, IDS_PREFS_THEMES_NOTE_DISK_THEME));
    }
    else
    {
        SetThemesNoteText(state, LoadStringResource(nullptr, IDS_PREFS_THEMES_NOTE_BUILTIN_THEME));
    }

    if (def)
    {
        SetThemesNameText(state, def->name);
    }
    else
    {
        SetThemesNameText(state, GetBuiltinThemeName(themeId));
    }

    ThemesPane::UpdateEditorFromSelection(host, state);
    RECT rc{};
    GetClientRect(host, &rc);
    PostMessageW(host, WM_SIZE, SIZE_RESTORED, MAKELPARAM(std::max(0l, rc.right - rc.left), std::max(0l, rc.bottom - rc.top)));
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

    applyD2D(L"viewer.diff.addedBackground", theme.viewerDiff.addedBackground);
    applyD2D(L"viewer.diff.removedBackground", theme.viewerDiff.removedBackground);
    applyD2D(L"viewer.diff.contextBackground", theme.viewerDiff.contextBackground);
    applyD2D(L"viewer.diff.headerBackground", theme.viewerDiff.headerBackground);
    applyD2D(L"viewer.diff.bannerBackground", theme.viewerDiff.bannerBackground);
    applyD2D(L"viewer.diff.placeholderBackground", theme.viewerDiff.placeholderBackground);
    applyD2D(L"viewer.diff.divider", theme.viewerDiff.divider);

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
        theme.metaPerf        = D2D1::ColorF(0.30f, 0.82f, 0.55f);
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
        theme.metaPerf        = D2D1::ColorF(0.40f, 0.90f, 0.62f);
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
        theme.metaPerf        = ColorFromCOLORREF(text, 1.0f);
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
        applyOverride(L"monitor.textView.metaPerf", theme.metaPerf);
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
    if (key == L"monitor.textView.metaPerf")
    {
        return ArgbFromD2DColorF(monitorTheme.metaPerf);
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
    if (state.pageHostWindow)
    {
        RedrawWindow(state.pageHostWindow, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
    }
    RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN | RDW_UPDATENOW);
    if (state.owner && IsWindow(state.owner) != FALSE)
    {
        PostMessageW(state.owner, WndMsg::kSettingsApplied, 0, 0);
    }
}

void PickThemeColorIntoEditor(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    uint32_t currentArgb           = 0xFF000000u;
    uint32_t alpha                 = 0xFFu;
    const std::wstring currentText = state.themesColorText;
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

    const uint32_t rgb  = (static_cast<uint32_t>(GetRValue(cc.rgbResult)) << 16) | (static_cast<uint32_t>(GetGValue(cc.rgbResult)) << 8) |
                          static_cast<uint32_t>(GetBValue(cc.rgbResult));
    const uint32_t argb = (alpha << 24) | rgb;

    std::wstring text;
    text                       = Common::Settings::FormatColor(argb);
    state.refreshingThemesPage = true;
    const auto reset           = wil::scope_exit([&] { state.refreshingThemesPage = false; });
    SetThemesColorText(state, text);
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

    std::wstring key = state.themesKeyText;
    key.erase(key.begin(), std::find_if(key.begin(), key.end(), [](wchar_t ch) noexcept { return ! std::iswspace(ch); }));
    key.erase(std::find_if(key.rbegin(), key.rend(), [](wchar_t ch) noexcept { return ! std::iswspace(ch); }).base(), key.end());
    if (! IsValidThemeColorKey(key))
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_WARNING, LoadStringResource(nullptr, IDS_CAPTION_WARNING), LoadStringResource(nullptr, IDS_PREFS_THEMES_WARNING_ENTER_COLOR_KEY));
        return;
    }

    const std::wstring valueText = state.themesColorText;
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

    std::wstring key = state.themesKeyText;
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

void ResetSelectedThemeToDefaults(HWND host, PreferencesDialogState& state) noexcept
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

    def->colors.clear();
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
    ApplyWindowChromeTheme(dlg, state.theme, WindowBackdropTarget::Tool, GetActiveWindow() == dlg);

    state.backgroundBrush.reset(CreateSolidBrush(state.theme.windowBackground));
    state.cardBackgroundColor = UiMetrics::GetControlSurfaceColor(state.theme);
    state.inputBrush.reset();
    state.inputFocusedBrush.reset();
    state.inputDisabledBrush.reset();
    state.cardBrush.reset();

    state.inputBackgroundColor         = UiMetrics::BlendColor(state.cardBackgroundColor, state.theme.windowBackground, state.theme.dark ? 50 : 30, 255);
    state.inputFocusedBackgroundColor  = UiMetrics::BlendColor(state.inputBackgroundColor, state.theme.menu.text, state.theme.dark ? 20 : 16, 255);
    state.inputDisabledBackgroundColor = UiMetrics::BlendColor(state.theme.windowBackground, state.inputBackgroundColor, state.theme.dark ? 70 : 40, 255);
    if (! state.theme.systemHighContrast)
    {
        state.cardBrush.reset(CreateSolidBrush(state.cardBackgroundColor));
        state.inputBrush.reset(CreateSolidBrush(state.inputBackgroundColor));
        state.inputFocusedBrush.reset(CreateSolidBrush(state.inputFocusedBackgroundColor));
        state.inputDisabledBrush.reset(CreateSolidBrush(state.inputDisabledBackgroundColor));
    }

    if (state.categoryTreeWindow)
    {
        if (state.categoryTreeUsesDxUi)
        {
            InvalidateRect(state.categoryTreeWindow, nullptr, TRUE);
        }
        SendMessageW(state.categoryTreeWindow, WM_THEMECHANGED, 0, 0);
        InvalidateRect(state.categoryTreeWindow, nullptr, TRUE);
    }
    if (state.pageHostWindow)
    {
        if (state.theme.systemHighContrast)
        {
            SetWindowTheme(state.pageHostWindow, L"", nullptr);
        }
        else
        {
            const wchar_t* hostTheme = state.theme.dark ? L"DarkMode_Explorer" : L"Explorer";
            SetWindowTheme(state.pageHostWindow, hostTheme, nullptr);
        }
        SendMessageW(state.pageHostWindow, WM_THEMECHANGED, 0, 0);
        InvalidateRect(state.pageHostWindow, nullptr, TRUE);
    }

    RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_FRAME | RDW_ERASE | RDW_ALLCHILDREN);
}

struct ThemesDxPage
{
    ThemesDxPage()                               = default;
    ThemesDxPage(const ThemesDxPage&)            = delete;
    ThemesDxPage& operator=(const ThemesDxPage&) = delete;
    ThemesDxPage(ThemesDxPage&&)                 = delete;
    ThemesDxPage& operator=(ThemesDxPage&&)      = delete;

    Label* themeLabel        = nullptr;
    ComboBox* themeCombo     = nullptr;
    Label* nameLabel         = nullptr;
    TextField* nameEdit      = nullptr;
    Label* baseLabel         = nullptr;
    ComboBox* baseCombo      = nullptr;
    Button* loadFromFile     = nullptr;
    Button* duplicateTheme   = nullptr;
    Button* resetTheme       = nullptr;
    Button* saveTheme        = nullptr;
    Button* applyTemporarily = nullptr;
    Label* note              = nullptr;
    Label* searchLabel       = nullptr;
    TextField* searchEdit    = nullptr;
    Grid* colorsListControl  = nullptr;
    std::unique_ptr<IDxGridModel> colorsListModelStorage;
    ThemesGridModel* colorsListModel = nullptr;
    Label* keyLabel                  = nullptr;
    TextField* keyEdit               = nullptr;
    Label* colorLabel                = nullptr;
    ColorSwatch* colorSwatch         = nullptr;
    TextField* colorEdit             = nullptr;
    Button* pickColor                = nullptr;
    Button* setOverride              = nullptr;
    Button* removeOverride           = nullptr;

    void Detach() noexcept
    {
        themeLabel        = nullptr;
        themeCombo        = nullptr;
        nameLabel         = nullptr;
        nameEdit          = nullptr;
        baseLabel         = nullptr;
        baseCombo         = nullptr;
        loadFromFile      = nullptr;
        duplicateTheme    = nullptr;
        resetTheme        = nullptr;
        saveTheme         = nullptr;
        applyTemporarily  = nullptr;
        note              = nullptr;
        searchLabel       = nullptr;
        searchEdit        = nullptr;
        colorsListControl = nullptr;
        colorsListModel   = nullptr;
        colorsListModelStorage.reset();
        keyLabel       = nullptr;
        keyEdit        = nullptr;
        colorLabel     = nullptr;
        colorSwatch    = nullptr;
        colorEdit      = nullptr;
        pickColor      = nullptr;
        setOverride    = nullptr;
        removeOverride = nullptr;
    }
};

struct ThemesPane::DxState
{
    DxState()                          = default;
    DxState(const DxState&)            = delete;
    DxState& operator=(const DxState&) = delete;
    DxState(DxState&&)                 = delete;
    DxState& operator=(DxState&&)      = delete;

    ThemesDxPage page;

    void Detach() noexcept
    {
        page.Detach();
    }
};

ThemesPane::ThemesPane() = default;

ThemesPane::~ThemesPane()
{
    DetachDxHosts();
}

bool ThemesPane::EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept
{
    _state           = &state;
    _hostWindow      = parent;
    _pageHostDx      = state.pageHostDxHost;
    _pageContentRoot = state.pageHostDxContentRootControl;
    if (! _pageHostDx || ! _pageContentRoot)
    {
        Debug::Error(L"Preferences.Themes: Shared page-host DX surface is unavailable; DxUi controls cannot be created.");
        return false;
    }

    if (! _rebuildDxOnNextShow && _dxState && PrefsUi::HasRetainedDxChildren(_pageContentRoot))
    {
        ApplyDxTheme(state);
        SyncDxControlsFromState(state);
        LogThemesDxState(
            L"ensure-dxhosts-reuse", _pageHost, _hostWindow, _pageHostDx, dynamic_cast<const Panel*>(_pageContentRoot), _dxState.get(), _rebuildDxOnNextShow);
        return true;
    }

    auto dxState = std::make_unique<DxState>();
    _pageHostDx->ResetInteractionState();
    _pageContentRoot->ClearChildren();
    auto* root = _pageContentRoot;

    dxState->page.themeLabel        = root->AddChild<Label>();
    dxState->page.themeCombo        = root->AddChild<ComboBox>();
    dxState->page.nameLabel         = root->AddChild<Label>();
    dxState->page.nameEdit          = root->AddChild<TextField>();
    dxState->page.baseLabel         = root->AddChild<Label>();
    dxState->page.baseCombo         = root->AddChild<ComboBox>();
    dxState->page.loadFromFile      = root->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_LOAD_FROM_FILE));
    dxState->page.duplicateTheme    = root->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_DUPLICATE));
    dxState->page.resetTheme        = root->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_RESET_DEFAULTS));
    dxState->page.saveTheme         = root->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_SAVE_THEME));
    dxState->page.applyTemporarily  = root->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_APPLY_TEMPORARILY));
    dxState->page.note              = root->AddChild<Label>();
    dxState->page.searchLabel       = root->AddChild<Label>();
    dxState->page.searchEdit        = root->AddChild<TextField>();
    dxState->page.colorsListControl = root->AddChild<Grid>();
    dxState->page.keyLabel          = root->AddChild<Label>();
    dxState->page.keyEdit           = root->AddChild<TextField>();
    dxState->page.colorLabel        = root->AddChild<Label>();
    dxState->page.colorSwatch       = root->AddChild<ColorSwatch>();
    dxState->page.colorEdit         = root->AddChild<TextField>();
    dxState->page.pickColor         = root->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_PICK));
    dxState->page.setOverride       = root->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_SET));
    dxState->page.removeOverride    = root->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_CLEAR));

    dxState->page.note->SetFontRole(FontRole::Small);
    dxState->page.note->SetMultiline(true);
    dxState->page.themeCombo->SetVariant(ComboBoxVariant::Window);
    dxState->page.baseCombo->SetVariant(ComboBoxVariant::Window);
    dxState->page.colorsListControl->SetDelegate(this);
    dxState->page.colorsListControl->SetSelectionMode(GridSelectionMode::Single);
    dxState->page.colorsListControl->SetHeaderHeightDip(30.0f);
    dxState->page.colorsListControl->SetRowHeightDip(30.0f);
    dxState->page.colorsListControl->SetLineClamp(1u);

    auto model                    = std::make_unique<ThemesGridModel>();
    dxState->page.colorsListModel = model.get();
    dxState->page.colorsListControl->SetModel(dxState->page.colorsListModel);
    dxState->page.colorsListModelStorage = std::move(model);

    dxState->page.themeCombo->SetOnSelectionChanged([this](const size_t selectedIndex) noexcept
    {
        if (_syncingDxInputs || ! _state || ! _hostWindow || IsWindow(_hostWindow) == FALSE || _state->refreshingThemesPage)
        {
            return;
        }

        if (selectedIndex >= _state->themeComboItems.size())
        {
            return;
        }

        const ThemeComboItem& selected = _state->themeComboItems[selectedIndex];
        _state->workingSettings.theme.currentThemeId.assign(selected.id);
        if (selected.source != ThemeSchemaSource::New)
        {
            SetDirty(GetParent(_hostWindow), *_state);
        }
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::ThemesThemeChanged));
    });

    dxState->page.baseCombo->SetOnSelectionChanged([this](const size_t selectedIndex) noexcept
    {
        if (_syncingDxInputs || ! _state || ! _hostWindow || IsWindow(_hostWindow) == FALSE || _state->refreshingThemesPage)
        {
            return;
        }

        auto* def = FindWorkingThemeDefinition(*_state, _state->workingSettings.theme.currentThemeId);
        if (! def)
        {
            return;
        }

        if (selectedIndex == 0u)
        {
            def->baseThemeId.assign(L"builtin/system");
        }
        else
        {
            const size_t optionIndex = selectedIndex - 1u;
            if (optionIndex >= kBuiltinThemeOptions.size())
            {
                return;
            }

            def->baseThemeId.assign(kBuiltinThemeOptions[optionIndex].id);
        }

        SetDirty(GetParent(_hostWindow), *_state);
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::ThemesBaseChanged));
    });

    dxState->page.nameEdit->SetOnTextChanged([this](std::wstring_view text) noexcept
    {
        if (_syncingDxInputs || ! _state || _state->refreshingThemesPage)
        {
            return;
        }

        _state->themesNameText.assign(text);

        auto* def = FindWorkingThemeDefinition(*_state, _state->workingSettings.theme.currentThemeId);
        if (! def || text.empty())
        {
            return;
        }

        def->name.assign(text);
        if (_hostWindow && IsWindow(_hostWindow) != FALSE)
        {
            SetDirty(GetParent(_hostWindow), *_state);
        }
    });
    dxState->page.nameEdit->SetOnBlur([this]() noexcept
    {
        if (_syncingDxInputs || ! _state || ! _hostWindow || IsWindow(_hostWindow) == FALSE)
        {
            return;
        }

        auto* dialogState = PrefsUi::GetDialogState(_hostWindow);
        if (! dialogState || dialogState->refreshingThemesPage)
        {
            return;
        }

        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::ThemesNameBlur));
    });

    dxState->page.searchEdit->SetOnTextChanged([this](std::wstring_view text) noexcept
    {
        if (_syncingDxInputs)
        {
            return;
        }

        if (_state)
        {
            _state->themesSearchText.assign(text);
        }

        if (! _state || _state->refreshingThemesPage || ! _hostWindow || IsWindow(_hostWindow) == FALSE)
        {
            return;
        }

        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::ThemesSearchChanged));
    });

    dxState->page.keyEdit->SetOnTextChanged([this](std::wstring_view text) noexcept
    {
        if (_syncingDxInputs || ! _state)
        {
            return;
        }

        _state->themesKeyText.assign(text);
    });

    dxState->page.colorEdit->SetOnTextChanged([this](std::wstring_view text) noexcept
    {
        if (_syncingDxInputs || ! _state)
        {
            return;
        }

        _state->themesColorText.assign(text);
        if (_dxState)
        {
            SyncDxSwatchFromState(*_state);
        }
    });

    const auto bindButton = [this](Button* button, const auto& callback) noexcept
    {
        if (! button)
        {
            return;
        }

        button->SetOnClick([this, callback]() noexcept
        {
            if (! _state || ! _hostWindow || IsWindow(_hostWindow) == FALSE)
            {
                return;
            }

            callback(_hostWindow, *_state);
            if (_dxState)
            {
                ApplyDxTheme(*_state);
                SyncDxControlsFromState(*_state);
            }
        });
    };

    bindButton(dxState->page.pickColor, PickThemeColorIntoEditor);
    bindButton(dxState->page.setOverride, SetThemeOverrideFromEditor);
    bindButton(dxState->page.removeOverride, ClearThemeOverrideFromEditor);
    bindButton(dxState->page.loadFromFile, LoadThemeFromFile);
    bindButton(dxState->page.duplicateTheme, DuplicateSelectedTheme);
    bindButton(dxState->page.resetTheme, ResetSelectedThemeToDefaults);
    bindButton(dxState->page.saveTheme, SaveThemeToFile);
    bindButton(dxState->page.applyTemporarily, ApplyThemeTemporarily);
    if (dxState->page.colorSwatch)
    {
        dxState->page.colorSwatch->SetOnClick([this]() noexcept
        {
            if (! _state || ! _hostWindow || IsWindow(_hostWindow) == FALSE)
            {
                return;
            }

            PickThemeColorIntoEditor(_hostWindow, *_state);
            if (_dxState)
            {
                ApplyDxTheme(*_state);
                SyncDxControlsFromState(*_state);
            }
        });
    }

    _dxState             = std::move(dxState);
    _rebuildDxOnNextShow = false;
    _state               = &state;
    _hostWindow          = parent;
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    LogThemesDxState(
        L"ensure-dxhosts-create", _pageHost, _hostWindow, _pageHostDx, dynamic_cast<const Panel*>(_pageContentRoot), _dxState.get(), _rebuildDxOnNextShow);
    return true;
}

void ThemesPane::DetachDxHosts() noexcept
{
    if (_pageContentRoot && _pageHostDx && _pageHost && IsWindow(_pageHost) != FALSE)
    {
        _pageHostDx->ResetInteractionState();
        _pageContentRoot->ClearChildren();
    }
    _pageHostDx      = nullptr;
    _pageContentRoot = nullptr;

    if (_dxState)
    {
        _dxState->Detach();
        _dxState.reset();
    }

    _syncingDxInputs     = false;
    _syncingDxSelection  = false;
    _rebuildDxOnNextShow = false;
    _state               = nullptr;
    _hostWindow          = nullptr;
}

void ThemesPane::ApplyDxTheme(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return;
    }

    _pageHostDx->SetTheme(PrefsUi::MakeDxPalette(state.theme));
}

void ThemesPane::SyncDxControlsFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxState)
    {
        return;
    }

    ThemesDxPage& page = _dxState->page;
    page.themeLabel->SetText(LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_THEME));
    page.nameLabel->SetText(LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_NAME));
    page.baseLabel->SetText(LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_BASE));
    page.searchLabel->SetText(LoadStringResource(nullptr, IDS_PREFS_COMMON_SEARCH));
    page.keyLabel->SetText(LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_KEY));
    page.colorLabel->SetText(LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_COLOR));
    page.themeLabel->SetMnemonicTarget(page.themeCombo);
    page.nameLabel->SetMnemonicTarget(page.nameEdit);
    page.baseLabel->SetMnemonicTarget(page.baseCombo);
    page.searchLabel->SetMnemonicTarget(page.searchEdit);
    page.keyLabel->SetMnemonicTarget(page.keyEdit);
    page.colorLabel->SetMnemonicTarget(page.colorEdit);
    page.note->SetText(state.themesNoteText);

    _syncingDxInputs          = true;
    auto clearSyncingDxInputs = wil::scope_exit([this]() noexcept { _syncingDxInputs = false; });

    const auto syncEdit = [](TextField* dxEdit, std::wstring_view text, const bool enabled) noexcept
    {
        if (! dxEdit)
        {
            return;
        }

        dxEdit->SetText(std::wstring(text));
        dxEdit->SetEnabled(enabled);
    };

    const auto syncThemeCombo = [&](ComboBox* dxCombo) noexcept
    {
        if (! dxCombo)
        {
            return;
        }

        std::vector<ComboBox::Item> items;
        items.reserve(state.themeComboItems.size());
        for (const auto& item : state.themeComboItems)
        {
            items.push_back(ComboBox::Item{item.id, item.displayName});
        }

        dxCombo->SetItems(std::move(items));

        std::optional<size_t> selectedIndex;
        const std::wstring_view desiredId = state.workingSettings.theme.currentThemeId;
        for (size_t index = 0; index < state.themeComboItems.size(); ++index)
        {
            if (std::wstring_view(state.themeComboItems[index].id) == desiredId)
            {
                selectedIndex = index;
                break;
            }
        }

        dxCombo->SetSelectedIndex(selectedIndex);
        dxCombo->SetEnabled(true);
    };

    const auto syncBaseCombo = [&](ComboBox* dxCombo) noexcept
    {
        if (! dxCombo)
        {
            return;
        }

        std::vector<ComboBox::Item> items;
        items.reserve(kBuiltinThemeOptions.size() + 1u);
        items.push_back(ComboBox::Item{L"builtin/system", LoadStringResource(nullptr, IDS_PREFS_THEMES_BASE_NONE)});
        for (const auto& option : kBuiltinThemeOptions)
        {
            const std::wstring name = GetBuiltinThemeName(option.id);
            items.push_back(ComboBox::Item{std::wstring(option.id), name.empty() ? std::wstring(option.id) : name});
        }

        dxCombo->SetItems(std::move(items));

        std::optional<size_t> selectedIndex = 0u;
        if (const auto* def = FindThemeDefinitionById(state.workingSettings.theme.themes, state.workingSettings.theme.currentThemeId))
        {
            if (! def->baseThemeId.empty())
            {
                for (size_t index = 0; index < kBuiltinThemeOptions.size(); ++index)
                {
                    if (kBuiltinThemeOptions[index].id == def->baseThemeId)
                    {
                        selectedIndex = index + 1u;
                        break;
                    }
                }
            }
        }

        dxCombo->SetSelectedIndex(selectedIndex);
        dxCombo->SetEnabled(FindThemeDefinitionById(state.workingSettings.theme.themes, state.workingSettings.theme.currentThemeId) != nullptr);
    };

    syncThemeCombo(page.themeCombo);
    syncBaseCombo(page.baseCombo);
    if (page.searchEdit)
    {
        page.searchEdit->SetText(state.themesSearchText);
    }

    if (page.colorsListControl && page.colorsListModel)
    {
        std::vector<ThemesGridRow> rows = BuildThemesColorRows(state);
        page.colorsListModel->SetRows(std::move(rows));

        _syncingDxSelection             = true;
        const std::wstring& selectedKey = state.themesSelectedColorKey;
        if (! selectedKey.empty())
        {
            const auto dxRowIndex = page.colorsListModel->FindRowIndexByKey(selectedKey);
            if (dxRowIndex.has_value())
            {
                page.colorsListControl->GetSelectionModel().SetSingle(page.colorsListModel->GetStableRowId(dxRowIndex.value()));
            }
            else
            {
                page.colorsListControl->GetSelectionModel().Clear();
            }
        }
        else if (! page.colorsListModel->GetRows().empty())
        {
            page.colorsListControl->GetSelectionModel().SetSingle(page.colorsListModel->GetStableRowId(0));
        }
        _syncingDxSelection = false;

        page.colorsListControl->SetEnabled(true);
        page.colorsListControl->NotifyDataChanged();
    }

    bool editable = false;
    if (const auto* def = FindThemeDefinitionForDisplay(state, state.workingSettings.theme.currentThemeId, editable))
    {
        UNREFERENCED_PARAMETER(def);
    }

    syncEdit(page.nameEdit, state.themesNameText, editable);
    syncEdit(page.keyEdit, state.themesKeyText, editable);
    syncEdit(page.colorEdit, state.themesColorText, editable);
    SyncDxSwatchFromState(state);

    const auto syncButton = [&](Button* dxButton, const UINT textId, const bool enabled) noexcept
    {
        if (! dxButton)
        {
            return;
        }

        dxButton->SetText(LoadStringResource(nullptr, textId));
        dxButton->SetEnabled(enabled);
    };

    syncButton(page.pickColor, IDS_PREFS_THEMES_BUTTON_PICK, editable);
    syncButton(page.setOverride, IDS_PREFS_THEMES_BUTTON_SET, editable);
    syncButton(page.removeOverride, IDS_PREFS_THEMES_BUTTON_CLEAR, editable);
    syncButton(page.loadFromFile, IDS_PREFS_THEMES_BUTTON_LOAD_FROM_FILE, true);
    syncButton(page.duplicateTheme, IDS_PREFS_THEMES_BUTTON_DUPLICATE, ! editable);
    syncButton(page.resetTheme, IDS_PREFS_KEYBOARD_BUTTON_RESET_DEFAULTS, editable);
    syncButton(page.saveTheme, IDS_PREFS_THEMES_BUTTON_SAVE_THEME, editable);
    syncButton(page.applyTemporarily, IDS_PREFS_THEMES_BUTTON_APPLY_TEMPORARILY, true);
    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void ThemesPane::SyncDxSwatchFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _dxState->page.colorSwatch)
    {
        return;
    }

    std::optional<uint32_t> swatchArgb;
    const std::wstring colorText = state.themesColorText;
    uint32_t parsedArgb          = 0u;
    if (! colorText.empty() && Common::Settings::TryParseColor(colorText, parsedArgb))
    {
        swatchArgb = parsedArgb;
    }

    _dxState->page.colorSwatch->SetSwatchValue(swatchArgb);
    bool editable = false;
    if (const auto* def = FindThemeDefinitionForDisplay(state, state.workingSettings.theme.currentThemeId, editable))
    {
        UNREFERENCED_PARAMETER(def);
    }
    _dxState->page.colorSwatch->SetEnabled(editable);
}

void ThemesPane::LayoutDxPage(HWND host,
                              const PreferencesDialogState& state,
                              int x,
                              int& y,
                              int width,
                              int margin,
                              int gapY,
                              int sectionY,
                              const PreferencesTypographyContext& typography) noexcept
{
    if (! host || ! _dxState || ! _pageHostDx || ! _pageContentRoot)
    {
        return;
    }

    ApplyDxTheme(state);
    SyncDxControlsFromState(state);

    using namespace PrefsLayoutConstants;

    Debug::Perf::Scope layoutPerf(L"preferences.ui.themes_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(static_cast<uint64_t>(typography.dpi));

    const UINT dpi        = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);
    const auto pxToDip    = [dpi](const int pixels) noexcept { return (static_cast<float>(pixels) * 96.0f) / static_cast<float>(dpi); };
    const int rowHeight   = std::max(1, UiMetrics::ScaleDip(dpi, kRowHeightDip));
    const int labelHeight = std::max(1, UiMetrics::ScaleDip(dpi, kTitleHeightDip));
    const int gapX        = UiMetrics::ScaleDip(dpi, kToggleGapXDip);

    RECT hostClient{};
    GetClientRect(host, &hostClient);
    const int hostBottom        = std::max(0l, hostClient.bottom - hostClient.top);
    const int hostContentBottom = std::max(0, hostBottom - margin);

    ThemesDxPage& page = _dxState->page;
    int localY         = y;

    const int themeLabelWidth      = std::min(width, UiMetrics::ScaleDip(dpi, 60));
    const int editWidth            = std::max(0, width - themeLabelWidth - gapX);
    const auto placeLabeledControl = [&](Label* label, Control* control, int controlWidth) noexcept
    {
        controlWidth       = std::max(0, std::min(editWidth, controlWidth));
        const int labelX   = x;
        const int controlX = x + themeLabelWidth + gapX;
        if (label)
        {
            label->SetBounds(D2D1::RectF(pxToDip(labelX),
                                         pxToDip(localY + (rowHeight - labelHeight) / 2),
                                         pxToDip(labelX + themeLabelWidth),
                                         pxToDip(localY + (rowHeight - labelHeight) / 2 + labelHeight)));
        }
        if (control)
        {
            control->SetBounds(D2D1::RectF(pxToDip(controlX), pxToDip(localY), pxToDip(controlX + controlWidth), pxToDip(localY + rowHeight)));
        }
        localY += rowHeight + gapY;
    };

    int themeWidth          = UiMetrics::ScaleDip(dpi, 220);
    const int minThemeWidth = UiMetrics::ScaleDip(dpi, 160);
    const int maxThemeWidth = std::max(minThemeWidth, std::max(0, editWidth));
    themeWidth              = std::clamp(themeWidth, minThemeWidth, maxThemeWidth);
    themeWidth              = std::min(themeWidth, UiMetrics::ScaleDip(dpi, 320));
    placeLabeledControl(page.themeLabel, page.themeCombo, themeWidth);

    placeLabeledControl(page.nameLabel, page.nameEdit, editWidth);

    int baseWidth = UiMetrics::ScaleDip(dpi, 180);
    baseWidth     = std::max(baseWidth, UiMetrics::ScaleDip(dpi, 100));
    placeLabeledControl(page.baseLabel, page.baseCombo, baseWidth);

    const int buttonHeight   = rowHeight;
    const int loadWidth      = std::min(width, UiMetrics::ScaleDip(dpi, 140));
    const int duplicateWidth = std::min(width, UiMetrics::ScaleDip(dpi, 110));
    const int resetWidth     = std::min(width, UiMetrics::ScaleDip(dpi, 140));
    const int saveWidth      = std::min(width, UiMetrics::ScaleDip(dpi, 120));
    const int applyWidth     = std::min(width, UiMetrics::ScaleDip(dpi, 150));

    int leftGroupWidth            = 0;
    const auto addLeftButtonWidth = [&](Button* button, const int buttonWidth) noexcept
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

    addLeftButtonWidth(page.loadFromFile, loadWidth);
    addLeftButtonWidth(page.duplicateTheme, duplicateWidth);
    addLeftButtonWidth(page.resetTheme, resetWidth);
    addLeftButtonWidth(page.saveTheme, saveWidth);

    const bool wrapApply = page.applyTemporarily && leftGroupWidth > 0 && (leftGroupWidth + gapX + applyWidth > width);
    const int row1Y      = localY;
    const int row2Y      = row1Y + buttonHeight + gapY;

    int leftButtonsX = x;
    if (page.loadFromFile)
    {
        page.loadFromFile->SetBounds(D2D1::RectF(pxToDip(leftButtonsX), pxToDip(row1Y), pxToDip(leftButtonsX + loadWidth), pxToDip(row1Y + buttonHeight)));
        leftButtonsX += loadWidth + gapX;
    }
    if (page.duplicateTheme)
    {
        page.duplicateTheme->SetBounds(
            D2D1::RectF(pxToDip(leftButtonsX), pxToDip(row1Y), pxToDip(leftButtonsX + duplicateWidth), pxToDip(row1Y + buttonHeight)));
        leftButtonsX += duplicateWidth + gapX;
    }
    if (page.resetTheme)
    {
        page.resetTheme->SetBounds(D2D1::RectF(pxToDip(leftButtonsX), pxToDip(row1Y), pxToDip(leftButtonsX + resetWidth), pxToDip(row1Y + buttonHeight)));
        leftButtonsX += resetWidth + gapX;
    }
    if (page.saveTheme)
    {
        page.saveTheme->SetBounds(D2D1::RectF(pxToDip(leftButtonsX), pxToDip(row1Y), pxToDip(leftButtonsX + saveWidth), pxToDip(row1Y + buttonHeight)));
    }
    if (page.applyTemporarily)
    {
        const int applyX = x + width - applyWidth;
        const int applyY = wrapApply ? row2Y : row1Y;
        page.applyTemporarily->SetBounds(D2D1::RectF(pxToDip(applyX), pxToDip(applyY), pxToDip(applyX + applyWidth), pxToDip(applyY + buttonHeight)));
    }

    localY = wrapApply ? (row2Y + buttonHeight + gapY) : (row1Y + buttonHeight + gapY);

    const std::wstring& noteText = state.themesNoteText;
    const int noteHeight         = noteText.empty() ? 0 : PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, width, noteText);
    if (page.note)
    {
        page.note->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(localY), pxToDip(x + width), pxToDip(localY + std::max(0, noteHeight))));
    }
    localY += std::max(0, noteHeight) + sectionY;

    const int searchLabelWidth = std::min(width, UiMetrics::ScaleDip(dpi, 52));
    const int searchEditWidth  = std::max(0, width - searchLabelWidth - gapX);
    if (page.searchLabel)
    {
        page.searchLabel->SetBounds(D2D1::RectF(pxToDip(x),
                                                pxToDip(localY + (rowHeight - labelHeight) / 2),
                                                pxToDip(x + searchLabelWidth),
                                                pxToDip(localY + (rowHeight - labelHeight) / 2 + labelHeight)));
    }
    if (page.searchEdit)
    {
        const int searchEditX = x + searchLabelWidth + gapX;
        page.searchEdit->SetBounds(D2D1::RectF(pxToDip(searchEditX), pxToDip(localY), pxToDip(searchEditX + searchEditWidth), pxToDip(localY + rowHeight)));
    }
    localY += rowHeight + gapY;

    const int editorHeight = rowHeight;
    const int editorTop    = std::max(localY, hostContentBottom - editorHeight);
    const int listTop      = localY;
    const int listBottom   = std::max(listTop, editorTop - gapY);
    const int listHeight   = std::max(0, listBottom - listTop);
    if (page.colorsListControl)
    {
        page.colorsListControl->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(listTop), pxToDip(x + width), pxToDip(listTop + listHeight)));
    }

    const int keyLabelWidth   = std::min(width, UiMetrics::ScaleDip(dpi, 34));
    const int colorLabelWidth = std::min(width, UiMetrics::ScaleDip(dpi, 44));
    const int pickWidth       = std::min(width, UiMetrics::ScaleDip(dpi, 70));
    const int setWidth        = std::min(width, UiMetrics::ScaleDip(dpi, 60));
    const int clearWidth      = std::min(width, UiMetrics::ScaleDip(dpi, 70));
    const int swatchWidth     = std::min(width, rowHeight);
    const int colorEditWidth  = std::min(width, UiMetrics::ScaleDip(dpi, 110));
    const int buttonsWidth    = pickWidth + gapX + setWidth + gapX + clearWidth;
    const int editAreaWidth =
        std::max(0, width - keyLabelWidth - gapX - colorLabelWidth - gapX - swatchWidth - gapX - colorEditWidth - gapX - buttonsWidth - gapX);

    if (page.keyLabel)
    {
        page.keyLabel->SetBounds(D2D1::RectF(pxToDip(x),
                                             pxToDip(editorTop + (rowHeight - labelHeight) / 2),
                                             pxToDip(x + keyLabelWidth),
                                             pxToDip(editorTop + (rowHeight - labelHeight) / 2 + labelHeight)));
    }
    const int keyEditX = x + keyLabelWidth + gapX;
    if (page.keyEdit)
    {
        page.keyEdit->SetBounds(D2D1::RectF(pxToDip(keyEditX), pxToDip(editorTop), pxToDip(keyEditX + editAreaWidth), pxToDip(editorTop + rowHeight)));
    }

    const int colorLabelX = keyEditX + editAreaWidth + gapX;
    if (page.colorLabel)
    {
        page.colorLabel->SetBounds(D2D1::RectF(pxToDip(colorLabelX),
                                               pxToDip(editorTop + (rowHeight - labelHeight) / 2),
                                               pxToDip(colorLabelX + colorLabelWidth),
                                               pxToDip(editorTop + (rowHeight - labelHeight) / 2 + labelHeight)));
    }

    const int colorSwatchX = colorLabelX + colorLabelWidth + gapX;
    if (page.colorSwatch)
    {
        page.colorSwatch->SetBounds(
            D2D1::RectF(pxToDip(colorSwatchX), pxToDip(editorTop), pxToDip(colorSwatchX + swatchWidth), pxToDip(editorTop + rowHeight)));
    }

    const int colorEditX = colorSwatchX + swatchWidth + gapX;
    if (page.colorEdit)
    {
        page.colorEdit->SetBounds(D2D1::RectF(pxToDip(colorEditX), pxToDip(editorTop), pxToDip(colorEditX + colorEditWidth), pxToDip(editorTop + rowHeight)));
    }

    int buttonX = colorEditX + colorEditWidth + gapX;
    if (page.pickColor)
    {
        page.pickColor->SetBounds(D2D1::RectF(pxToDip(buttonX), pxToDip(editorTop), pxToDip(buttonX + pickWidth), pxToDip(editorTop + rowHeight)));
        buttonX += pickWidth + gapX;
    }
    if (page.setOverride)
    {
        page.setOverride->SetBounds(D2D1::RectF(pxToDip(buttonX), pxToDip(editorTop), pxToDip(buttonX + setWidth), pxToDip(editorTop + rowHeight)));
        buttonX += setWidth + gapX;
    }
    if (page.removeOverride)
    {
        page.removeOverride->SetBounds(D2D1::RectF(pxToDip(buttonX), pxToDip(editorTop), pxToDip(buttonX + clearWidth), pxToDip(editorTop + rowHeight)));
    }

    _pageHostDx->Invalidate();
    y = hostContentBottom;
}

void ThemesPane::OnVisibilityChanged(bool visible) noexcept
{
    if (! visible)
    {
        if (_pageHostDx)
        {
            _pageHostDx->ResetInteractionState();
        }
    }
}

void ThemesPane::Destroy(PreferencesDialogState& state) noexcept
{
    DetachDxHosts();
    state.themesNoteText.clear();
    state.themesNameText.clear();
    state.themesKeyText.clear();
    state.themesColorText.clear();
    state.themeComboItems.clear();
    state.refreshingThemesPage = false;
    _pageHost                  = nullptr;
    _pageHostDx                = nullptr;
    _pageContentRoot           = nullptr;
}

void ThemesPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageHost = parent;

    if (state.currentCategory != PrefCategory::Themes)
    {
        return;
    }

    if (! EnsureDxHosts(parent, state))
    {
        Debug::Error(L"Preferences.Themes: Failed to initialize DxUi hosts in CreateControls.");
        DetachDxHosts();
        return;
    }
}

void ThemesPane::LayoutPage(HWND host,
                            PreferencesDialogState& state,
                            int x,
                            int& y,
                            int width,
                            int margin,
                            int gapY,
                            int sectionY,
                            const PreferencesTypographyContext& typography) noexcept
{
    if (! host)
    {
        return;
    }

    if (EnsureDxHosts(_pageHost ? _pageHost : host, state))
    {
        LayoutDxPage(host, state, x, y, width, margin, gapY, sectionY, typography);
        return;
    }

    Debug::Error(L"Preferences.Themes: DxUi surface initialization failed in LayoutControls; page will not render correctly.");
}

void ThemesPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    _hostWindow = host;
    _state      = &state;
    RefreshThemesPage(host, state);

    if (EnsureDxHosts(_pageHost ? _pageHost : host, state))
    {
        ApplyDxTheme(state);
        SyncDxControlsFromState(state);
        LogThemesDxState(
            L"refresh-complete", _pageHost, _hostWindow, _pageHostDx, dynamic_cast<const Panel*>(_pageContentRoot), _dxState.get(), _rebuildDxOnNextShow);
    }
    else
    {
        Debug::Error(L"Preferences.Themes: Failed to ensure DxUi hosts during Refresh.");
    }
}

bool ThemesPane::HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept
{
    _hostWindow = host;
    _state      = &state;

    if (! _dxState)
    {
        return false;
    }

    const auto handled = [&]() noexcept
    {
        if (_dxState)
        {
            ApplyDxTheme(state);
            SyncDxControlsFromState(state);
        }
        return true;
    };

#pragma warning(suppress : 4061) // Not all enum values handled explicitly -- intentional; this pane only handles its own actions.
    switch (action)
    {
        case PreferencesDeferredActionKind::ThemesSearchChanged: RefreshThemesPage(host, state); return handled();

        case PreferencesDeferredActionKind::ThemesThemeChanged:
        {
            if (state.refreshingThemesPage)
            {
                return handled();
            }

            const ThemeComboItem* selected = TryGetSelectedThemeComboItem(state);
            if (! selected)
            {
                return handled();
            }

            if (selected->source == ThemeSchemaSource::New)
            {
                BeginNewThemeCreation(host, state);
                return handled();
            }

            state.workingSettings.theme.currentThemeId.assign(selected->id);
            SetDirty(GetParent(host), state);
            RefreshThemesPage(host, state);
            return handled();
        }

        case PreferencesDeferredActionKind::ThemesBaseChanged:
            if (state.refreshingThemesPage)
            {
                return handled();
            }

            RefreshThemesPage(host, state);
            return handled();

        case PreferencesDeferredActionKind::ThemesNameBlur:
            if (state.refreshingThemesPage)
            {
                return handled();
            }

            SyncSelectedUserThemeIdToName(host, state);
            return handled();
        case PreferencesDeferredActionKind::ViewersSearchChanged:
        case PreferencesDeferredActionKind::KeyboardSearchChanged:
        case PreferencesDeferredActionKind::KeyboardScopeChanged:
        case PreferencesDeferredActionKind::KeyboardAssign:
        case PreferencesDeferredActionKind::KeyboardRemove:
        case PreferencesDeferredActionKind::KeyboardReset:
        case PreferencesDeferredActionKind::KeyboardImport:
        case PreferencesDeferredActionKind::KeyboardExport:
        case PreferencesDeferredActionKind::PluginsSearchChanged:
        case PreferencesDeferredActionKind::PluginsConfigure:
        case PreferencesDeferredActionKind::PluginsTest:
        case PreferencesDeferredActionKind::PluginsTestAll:
        case PreferencesDeferredActionKind::FileOperationsBandwidthPresetChanged:
        case PreferencesDeferredActionKind::CompareDirectoriesIgnoreToggleChanged: return false;
        default: return false;
    }
}

void ThemesPane::UpdateEditorFromSelection(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    const std::wstring selectedKey = state.themesSelectedColorKey;
    if (selectedKey.empty())
    {
        state.refreshingThemesPage = true;
        const auto reset           = wil::scope_exit([&] { state.refreshingThemesPage = false; });
        SetThemesKeyText(state, L"");
        SetThemesColorText(state, L"");
        return;
    }

    std::wstring valueText;
    const auto themeIdOpt = TryGetSelectedThemeId(state);
    if (themeIdOpt.has_value())
    {
        bool editable                       = false;
        const auto* def                     = FindThemeDefinitionForDisplay(state, themeIdOpt.value(), editable);
        const std::wstring_view baseThemeId = (def && ! def->baseThemeId.empty()) ? std::wstring_view(def->baseThemeId) : themeIdOpt.value();
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

        const auto colorOpt = TryGetEffectiveThemeColorArgb(appTheme, monitorTheme, overrides, selectedKey);
        if (colorOpt.has_value())
        {
            valueText = Common::Settings::FormatColor(colorOpt.value());
        }
    }

    state.refreshingThemesPage = true;
    const auto reset           = wil::scope_exit([&] { state.refreshingThemesPage = false; });
    SetThemesKeyText(state, selectedKey);
    SetThemesColorText(state, valueText);
}

void ThemesPane::OnGridSelectionChanged()
{
    if (! _state || ! _dxState || ! _dxState->page.colorsListControl || ! _dxState->page.colorsListModel || _syncingDxSelection)
    {
        return;
    }

    const auto selectedRowIds = _dxState->page.colorsListControl->GetSelectionModel().GetOrderedSelection();
    std::wstring selectedKey;
    if (! selectedRowIds.empty())
    {
        const auto& rows = _dxState->page.colorsListModel->GetRows();
        const auto it    = std::find_if(rows.begin(), rows.end(), [&](const ThemesGridRow& row) noexcept { return row.stableId == selectedRowIds.front(); });
        if (it != rows.end())
        {
            selectedKey = it->key;
        }
    }

    _state->themesSelectedColorKey = selectedKey;

    if (_hostWindow)
    {
        UpdateEditorFromSelection(_hostWindow, *_state);
    }
    if (_dxState)
    {
        SyncDxControlsFromState(*_state);
    }
}

#ifdef ENABLE_TESTS
size_t ThemesPane::DebugListRowCount() const noexcept
{
    if (! _dxState || ! _dxState->page.colorsListModel)
    {
        return 0u;
    }

    return _dxState->page.colorsListModel->GetRowCount();
}

RedSalamander::DxUi::GridVisibleWorkMetrics ThemesPane::DebugListVisibleWorkMetrics() const noexcept
{
    if (! _dxState || ! _dxState->page.colorsListControl)
    {
        return {};
    }

    return _dxState->page.colorsListControl->GetVisibleWorkMetrics();
}

uint64_t ThemesPane::DebugListRenderCount() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return 0u;
    }

#ifdef ENABLE_TESTS
    return _pageHostDx->DebugGetRenderCount();
#else
    return 0u;
#endif
}

uint64_t ThemesPane::DebugListResizeCount() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return 0u;
    }

#ifdef ENABLE_TESTS
    return _pageHostDx->DebugGetResizeCount();
#else
    return 0u;
#endif
}

uint64_t ThemesPane::DebugListResizeFailureCount() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return 0u;
    }

#ifdef ENABLE_TESTS
    return _pageHostDx->DebugGetResizeFailureCount();
#else
    return 0u;
#endif
}

PreferencesThemesDebugFocusTarget ThemesPane::DebugGetFocusTarget() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return PreferencesThemesDebugFocusTarget::None;
    }

    RedSalamander::DxUi::Control* const focusedControl = _pageHostDx->GetFocusControl();
    if (! focusedControl)
    {
        return PreferencesThemesDebugFocusTarget::None;
    }

    const auto& page = _dxState->page;
    if (focusedControl == page.themeCombo)
        return PreferencesThemesDebugFocusTarget::ThemeCombo;
    if (focusedControl == page.nameEdit)
        return PreferencesThemesDebugFocusTarget::NameField;
    if (focusedControl == page.baseCombo)
        return PreferencesThemesDebugFocusTarget::BaseCombo;
    if (focusedControl == page.loadFromFile)
        return PreferencesThemesDebugFocusTarget::LoadFromFileButton;
    if (focusedControl == page.duplicateTheme)
        return PreferencesThemesDebugFocusTarget::DuplicateButton;
    if (focusedControl == page.resetTheme)
        return PreferencesThemesDebugFocusTarget::ResetButton;
    if (focusedControl == page.saveTheme)
        return PreferencesThemesDebugFocusTarget::SaveButton;
    if (focusedControl == page.applyTemporarily)
        return PreferencesThemesDebugFocusTarget::ApplyTemporarilyButton;
    if (focusedControl == page.searchEdit)
        return PreferencesThemesDebugFocusTarget::SearchField;
    if (focusedControl == page.colorsListControl)
        return PreferencesThemesDebugFocusTarget::ColorsGrid;
    if (focusedControl == page.keyEdit)
        return PreferencesThemesDebugFocusTarget::KeyField;
    if (focusedControl == page.colorEdit)
        return PreferencesThemesDebugFocusTarget::ColorField;
    if (focusedControl == page.pickColor)
        return PreferencesThemesDebugFocusTarget::PickButton;
    if (focusedControl == page.setOverride)
        return PreferencesThemesDebugFocusTarget::SetButton;
    if (focusedControl == page.removeOverride)
        return PreferencesThemesDebugFocusTarget::ClearButton;

    return PreferencesThemesDebugFocusTarget::None;
}

bool ThemesPane::DebugGetListRowClientRect(const size_t rowIndex, RECT& outRect) const noexcept
{
    if (! _dxState || ! _dxState->page.colorsListControl || ! _pageHostDx)
    {
        return false;
    }

    const auto rowRect = _dxState->page.colorsListControl->GetVisibleRowRect(rowIndex);
    if (! rowRect.has_value())
    {
        return false;
    }

    outRect.left   = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(rowRect->left)));
    outRect.top    = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(rowRect->top)));
    outRect.right  = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(rowRect->right)));
    outRect.bottom = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(rowRect->bottom)));
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

bool ThemesPane::DebugGetListHeaderClientRect(const size_t columnIndex, RECT& outRect) const noexcept
{
    if (! _dxState || ! _dxState->page.colorsListControl || ! _dxState->page.colorsListModel || ! _pageHostDx ||
        columnIndex >= _dxState->page.colorsListModel->GetColumnCount())
    {
        return false;
    }

    const auto headerRect = _dxState->page.colorsListControl->GetVisibleColumnHeaderRect(columnIndex);
    if (! headerRect.has_value())
    {
        return false;
    }

    outRect.left   = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->left)));
    outRect.top    = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->top)));
    outRect.right  = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->right)));
    outRect.bottom = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->bottom)));
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

bool ThemesPane::DebugSelectListRow(const size_t rowIndex) noexcept
{
    if (! _dxState || ! _dxState->page.colorsListControl || ! _dxState->page.colorsListModel)
    {
        return false;
    }

    const auto& rows = _dxState->page.colorsListModel->GetRows();
    if (rowIndex >= rows.size())
    {
        return false;
    }

    _dxState->page.colorsListControl->GetSelectionModel().SetSingle(rows[rowIndex].stableId);
    OnGridSelectionChanged();
    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
    return true;
}

bool ThemesPane::DebugSetSearchText(std::wstring_view text) noexcept
{
    if (! _state)
    {
        return false;
    }

    _state->themesSearchText.assign(text);
    if (_dxState && _dxState->page.searchEdit)
    {
        _dxState->page.searchEdit->SetText(std::wstring(text));
    }

    // TextField::SetText does not fire the OnTextChanged callback, so the
    // normal callback → Refresh chain is not triggered.
    // Manually refresh to re-filter the grid with the new search text.
    if (_hostWindow && _state)
    {
        Refresh(_hostWindow, *_state);
    }

    return _dxState && _dxState->page.searchEdit;
}

bool ThemesPane::DebugFocusSearchField() noexcept
{
    if (! _dxState || ! _dxState->page.searchEdit || ! _pageHostDx)
    {
        return false;
    }

    _pageHostDx->SetFocusControl(_dxState->page.searchEdit);
    return true;
}

bool ThemesPane::DebugScrollListByWheelDetents(const int detents) noexcept
{
    if (detents == 0 || ! _dxState || ! _dxState->page.colorsListControl || ! _pageHostDx)
    {
        return false;
    }

    _pageHostDx->SetFocusControl(_dxState->page.colorsListControl);

    const int direction = detents < 0 ? -1 : 1;
    const int steps     = std::abs(detents);
    for (int index = 0; index < steps; ++index)
    {
        const float wheelDelta = static_cast<float>(direction * WHEEL_DELTA);
        _dxState->page.colorsListControl->OnMouseWheel(*_pageHostDx, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0u);
    }

    return true;
}
#endif

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
    uint32_t parsed = 0;
    if (! state.themesColorText.empty() && Common::Settings::TryParseColor(state.themesColorText, parsed))
    {
        argb = parsed;
    }

    RECT swatch = dis->rcItem;
    InflateRect(&swatch, -UiMetrics::ScaleDip(dpi, PrefsLayoutConstants::kFramePaddingDip), -UiMetrics::ScaleDip(dpi, PrefsLayoutConstants::kFramePaddingDip));
    DrawRoundedColorSwatch(dis->hDC, swatch, dpi, state.theme, bg, argb, IsWindowEnabled(dis->hwndItem) != FALSE);
    return 1;
}

#ifdef ENABLE_TESTS
bool DebugSetPreferencesThemesNextBrowsePath(const std::wstring_view path) noexcept
{
    return DebugSetPreferencesThemesNextBrowsePathImpl(path);
}

bool DebugCancelPreferencesThemesNextBrowse() noexcept
{
#ifdef ENABLE_TESTS
    return DebugCancelPreferencesThemesNextBrowseImpl();
#else
    return false;
#endif
}
#endif
