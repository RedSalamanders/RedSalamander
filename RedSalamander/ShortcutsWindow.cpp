#include "Framework.h"

#include "ShortcutsWindow.h"

#include <algorithm>
#include <array>
#include <cwchar>
#include <cwctype>
#include <optional>
#include <unordered_map>
#include <vector>

#include <CommCtrl.h>
#include <Uxtheme.h>

#include "CommandRegistry.h"
#include "Helpers.h"
#include "SettingsSave.h"
#include "ShortcutManager.h"
#include "ShortcutText.h"
#include "ThemedControls.h"
#include "ThemedInputFrames.h"
#include "WindowMaximizeBehavior.h"
#include "WindowPlacementPersistence.h"
#include "resource.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

namespace
{
constexpr int kListCtrlId                = 100;
constexpr int kSearchEditId              = 101;
constexpr UINT_PTR kListHeaderSubclassId = 1u;
constexpr wchar_t kShortcutsWindowId[]   = L"ShortcutsWindow";
constexpr wchar_t kSettingsAppId[]       = L"RedSalamander";

constexpr int kGroupFunctionBar = 1;
constexpr int kGroupFolderView  = 2;

[[nodiscard]] std::wstring GetWindowTextString(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return {};
    }

    const int len = GetWindowTextLengthW(hwnd);
    if (len <= 0)
    {
        return {};
    }

    std::wstring text;
    text.resize(static_cast<size_t>(len));
    const int copied = GetWindowTextW(hwnd, text.data(), len + 1);
    if (copied <= 0)
    {
        return {};
    }

    if (static_cast<size_t>(copied) < text.size())
    {
        text.resize(static_cast<size_t>(copied));
    }

    return text;
}

[[nodiscard]] std::wstring_view TrimWhitespace(std::wstring_view text) noexcept
{
    while (! text.empty() && std::iswspace(static_cast<wint_t>(text.front())) != 0)
    {
        text.remove_prefix(1);
    }

    while (! text.empty() && std::iswspace(static_cast<wint_t>(text.back())) != 0)
    {
        text.remove_suffix(1);
    }

    return text;
}

[[nodiscard]] bool ContainsNoCase(std::wstring_view haystack, std::wstring_view needle) noexcept
{
    if (needle.empty())
    {
        return true;
    }

    if (needle.size() > haystack.size())
    {
        return false;
    }

    const auto it = std::search(haystack.begin(),
                                haystack.end(),
                                needle.begin(),
                                needle.end(),
                                [](wchar_t a, wchar_t b) noexcept { return std::towupper(static_cast<wint_t>(a)) == std::towupper(static_cast<wint_t>(b)); });

    return it != haystack.end();
}

[[nodiscard]] std::vector<std::pair<size_t, size_t>> FindAllMatchesNoCase(std::wstring_view text, std::wstring_view needle) noexcept
{
    std::vector<std::pair<size_t, size_t>> matches;
    if (needle.empty() || text.empty() || needle.size() > text.size())
    {
        return matches;
    }

    size_t startIndex = 0;
    while (startIndex + needle.size() <= text.size())
    {
        const auto it =
            std::search(text.begin() + static_cast<ptrdiff_t>(startIndex),
                        text.end(),
                        needle.begin(),
                        needle.end(),
                        [](wchar_t a, wchar_t b) noexcept { return std::towupper(static_cast<wint_t>(a)) == std::towupper(static_cast<wint_t>(b)); });

        if (it == text.end())
        {
            break;
        }

        const size_t index = static_cast<size_t>(it - text.begin());
        matches.emplace_back(index, needle.size());
        startIndex = index + needle.size();
    }

    return matches;
}

void DrawTextWithHighlights(HDC hdc,
                            std::wstring_view text,
                            const RECT& rc,
                            UINT format,
                            std::wstring_view query,
                            COLORREF textColor,
                            COLORREF highlightTextColor,
                            HBRUSH highlightBrush) noexcept
{
    if (! hdc || rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return;
    }

    query = TrimWhitespace(query);
    if (query.empty() || highlightBrush == nullptr || ! ContainsNoCase(text, query))
    {
        SetTextColor(hdc, textColor);
        DrawTextW(hdc, text.data(), static_cast<int>(text.size()), const_cast<RECT*>(&rc), format);
        return;
    }

    const std::vector<std::pair<size_t, size_t>> matches = FindAllMatchesNoCase(text, query);
    if (matches.empty())
    {
        SetTextColor(hdc, textColor);
        DrawTextW(hdc, text.data(), static_cast<int>(text.size()), const_cast<RECT*>(&rc), format);
        return;
    }

    TEXTMETRICW tm{};
    GetTextMetricsW(hdc, &tm);
    const int lineHeight = std::max(1, static_cast<int>(tm.tmHeight + tm.tmExternalLeading));

    int baseY = rc.top;
    if ((format & DT_VCENTER) != 0)
    {
        const int height = std::max(0l, rc.bottom - rc.top);
        baseY            = rc.top + std::max(0, (height - lineHeight) / 2);
    }

    int baseX = rc.left;
    if ((format & DT_RIGHT) != 0)
    {
        SIZE totalSize{};
        if (GetTextExtentPoint32W(hdc, text.data(), static_cast<int>(text.size()), &totalSize) != FALSE)
        {
            const int maxW   = std::max(0, static_cast<int>(rc.right - rc.left));
            const int totalW = std::max(0, static_cast<int>(totalSize.cx));
            const int w      = std::min(maxW, totalW);
            baseX            = static_cast<int>(rc.right) - w;
        }

        baseX = std::clamp(baseX, static_cast<int>(rc.left), static_cast<int>(rc.right));
    }

    const int saved = SaveDC(hdc);
    IntersectClipRect(hdc, rc.left, rc.top, rc.right, rc.bottom);

    RECT highlightRc{};
    highlightRc.top    = std::clamp(baseY, static_cast<int>(rc.top), static_cast<int>(rc.bottom));
    highlightRc.bottom = std::clamp(baseY + lineHeight, static_cast<int>(rc.top), static_cast<int>(rc.bottom));

    for (const auto& [index, length] : matches)
    {
        if (length == 0 || index >= text.size())
        {
            continue;
        }

        SIZE prefixSize{};
        if (index > 0)
        {
            GetTextExtentPoint32W(hdc, text.data(), static_cast<int>(index), &prefixSize);
        }

        const size_t clampedLength = std::min(length, text.size() - index);
        SIZE matchSize{};
        GetTextExtentPoint32W(hdc, text.data() + index, static_cast<int>(clampedLength), &matchSize);

        const int x0      = baseX + prefixSize.cx;
        const int x1      = x0 + matchSize.cx;
        highlightRc.left  = std::clamp(x0, static_cast<int>(rc.left), static_cast<int>(rc.right));
        highlightRc.right = std::clamp(x1, static_cast<int>(rc.left), static_cast<int>(rc.right));

        if (highlightRc.right > highlightRc.left && highlightRc.bottom > highlightRc.top)
        {
            FillRect(hdc, &highlightRc, highlightBrush);
        }
    }

    SetTextColor(hdc, textColor);
    DrawTextW(hdc, text.data(), static_cast<int>(text.size()), const_cast<RECT*>(&rc), format);

    SetTextColor(hdc, highlightTextColor);
    for (const auto& [index, length] : matches)
    {
        if (length == 0 || index >= text.size())
        {
            continue;
        }

        SIZE prefixSize{};
        if (index > 0)
        {
            GetTextExtentPoint32W(hdc, text.data(), static_cast<int>(index), &prefixSize);
        }

        const size_t clampedLength = std::min(length, text.size() - index);
        const int x                = baseX + prefixSize.cx;

        ExtTextOutW(hdc, x, baseY, 0, nullptr, text.data() + index, static_cast<UINT>(clampedLength), nullptr);
    }

    RestoreDC(hdc, saved);
}

[[nodiscard]] std::wstring GetCommandDisplayName(std::wstring_view commandId) noexcept
{
    return ShortcutText::GetCommandDisplayName(commandId);
}

[[nodiscard]] std::wstring GetCommandDescription(std::wstring_view commandId) noexcept
{
    const std::optional<unsigned int> descIdOpt = TryGetCommandDescriptionStringId(commandId);
    if (descIdOpt.has_value())
    {
        const std::wstring desc = LoadStringResource(nullptr, descIdOpt.value());
        if (! desc.empty())
        {
            return desc;
        }
    }

    return {};
}

[[nodiscard]] std::wstring FormatChordText(uint32_t vk, uint32_t modifiers) noexcept
{
    std::vector<std::wstring> parts;
    parts.reserve(4);

    if ((modifiers & ShortcutManager::kModCtrl) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_CTRL));
    }

    if ((modifiers & ShortcutManager::kModAlt) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_ALT));
    }

    if ((modifiers & ShortcutManager::kModShift) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_SHIFT));
    }

    parts.push_back(ShortcutText::VkToDisplayText(vk));

    std::wstring result;
    for (size_t i = 0; i < parts.size(); ++i)
    {
        if (parts[i].empty())
        {
            continue;
        }

        if (! result.empty())
        {
            result.append(L" + ");
        }
        result.append(parts[i]);
    }
    return result;
}

[[nodiscard]] bool IsConflictChord(uint32_t chordKey, const std::vector<uint32_t>& conflicts) noexcept
{
    return std::binary_search(conflicts.begin(), conflicts.end(), chordKey);
}

[[nodiscard]] COLORREF BlendColor(COLORREF base, COLORREF overlay, int overlayWeight, int denom) noexcept
{
    if (denom <= 0)
    {
        return base;
    }

    overlayWeight        = std::clamp(overlayWeight, 0, denom);
    const int baseWeight = denom - overlayWeight;

    const int r = (static_cast<int>(GetRValue(base)) * baseWeight + static_cast<int>(GetRValue(overlay)) * overlayWeight) / denom;
    const int g = (static_cast<int>(GetGValue(base)) * baseWeight + static_cast<int>(GetGValue(overlay)) * overlayWeight) / denom;
    const int b = (static_cast<int>(GetBValue(base)) * baseWeight + static_cast<int>(GetBValue(overlay)) * overlayWeight) / denom;
    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}

struct ShortcutRow final
{
    Common::Settings::ShortcutBinding binding;
    std::wstring displayName;
    std::wstring description;
    std::wstring keyText;
    uint32_t chordKey = 0;
    int groupId       = 0;
    bool conflict     = false;
    std::wstring conflictWith;
};

class ShortcutsWindow final
{
public:
    ShortcutsWindow() = default;
    ~ShortcutsWindow()
    {
        Destroy();
    }

    ShortcutsWindow(const ShortcutsWindow&)            = delete;
    ShortcutsWindow& operator=(const ShortcutsWindow&) = delete;
    ShortcutsWindow(ShortcutsWindow&&)                 = delete;
    ShortcutsWindow& operator=(ShortcutsWindow&&)      = delete;

    HWND Create(HWND owner,
                Common::Settings::Settings& settings,
                const Common::Settings::ShortcutsSettings& shortcuts,
                const ShortcutManager& shortcutManager,
                const AppTheme& theme) noexcept;

    void UpdateTheme(const AppTheme& theme) noexcept;
    void UpdateData(const Common::Settings::ShortcutsSettings& shortcuts, const ShortcutManager& shortcutManager) noexcept;

    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _hWnd.get();
    }

private:
    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static constexpr PCWSTR kClassName = L"RedSalamander.ShortcutsWindow";

    static LRESULT CALLBACK WndProcThunk(HWND hWindow, UINT msg, WPARAM wp, LPARAM lp);
    static LRESULT CALLBACK HeaderSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);

    void Destroy() noexcept;
    void OnCreate(HWND hwnd) noexcept;
    void OnDestroy() noexcept;
    void OnSize(UINT width, UINT height) noexcept;
    void OnPaint(HWND hwnd) noexcept;
    void OnActivate(HWND hwnd) noexcept;
    void OnHeaderPaint(HWND header) noexcept;
    LRESULT OnHeaderNcDestroy(HWND header, WPARAM wParam, LPARAM lParam, UINT_PTR subclassId) noexcept;
    LRESULT OnMeasureItem(MEASUREITEMSTRUCT* mis) noexcept;
    LRESULT OnDrawItem(DRAWITEMSTRUCT* dis) noexcept;
    void OnCommandMessage(UINT controlId, UINT notifyCode) noexcept;
    LRESULT OnNotify(const NMHDR* header, LPARAM lp) noexcept;
    LRESULT OnCustomDraw(NMLVCUSTOMDRAW* cd) noexcept;
    LRESULT OnGetInfoTip(NMLVGETINFOTIPW* tip) noexcept;
    LRESULT OnDpiChangedMessage(HWND hwnd, UINT dpi, const RECT* suggested) noexcept;
    LRESULT OnGetMinMaxInfoMessage(HWND hwnd, MINMAXINFO* info) noexcept;
    LRESULT OnNcDestroyMessage() noexcept;
    LRESULT OnCtlColorEdit(HDC hdc, HWND control) noexcept;
    void OnSearchChanged() noexcept;
    void EnsureFonts(UINT dpi) noexcept;

    void EnsureSearchControls(HWND hwnd) noexcept;
    void EnsureListView(HWND hwnd) noexcept;
    void EnsureColumns(UINT dpi) noexcept;
    void EnsureGroups() noexcept;
    void ApplyListTheme() noexcept;
    void PopulateList() noexcept;
    void ResizeWindowToContent(HWND hwnd) noexcept;
    void AutoSizeColumnsToContent(UINT dpi) noexcept;
    [[nodiscard]] int GetRowHeightPx(HDC hdc) const noexcept;

private:
    wil::unique_hwnd _hWnd;
    HINSTANCE _hInstance = nullptr;
    wil::unique_hwnd _searchFrame;
    HWND _searchEdit = nullptr;
    HWND _list       = nullptr;

    wil::unique_any<HIMAGELIST, decltype(&::ImageList_Destroy), ::ImageList_Destroy> _imageList;
    AppTheme _theme;

    std::wstring _searchQuery;

    ThemedInputFrames::FrameStyle _searchFrameStyle{};
    COLORREF _searchInputBackgroundColor         = RGB(255, 255, 255);
    COLORREF _searchInputFocusedBackgroundColor  = RGB(255, 255, 255);
    COLORREF _searchInputDisabledBackgroundColor = RGB(255, 255, 255);
    wil::unique_hbrush _searchInputBrush;
    wil::unique_hbrush _searchInputFocusedBrush;
    wil::unique_hbrush _searchInputDisabledBrush;

    Common::Settings::ShortcutsSettings _shortcuts;
    const ShortcutManager* _shortcutManager = nullptr;
    std::vector<ShortcutRow> _rows;

    Common::Settings::Settings* _settings = nullptr;

    wil::unique_hbrush _backgroundBrush;
    UINT _dpi = USER_DEFAULT_SCREEN_DPI;
    wil::unique_hfont _uiFont;
};

ShortcutsWindow* g_shortcutsWindow = nullptr;

ATOM ShortcutsWindow::RegisterWndClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

HWND ShortcutsWindow::Create(HWND owner,
                             Common::Settings::Settings& settings,
                             const Common::Settings::ShortcutsSettings& shortcuts,
                             const ShortcutManager& shortcutManager,
                             const AppTheme& theme) noexcept
{
    _hInstance = GetModuleHandleW(nullptr);
    if (! RegisterWndClass(_hInstance))
    {
        return nullptr;
    }

    _settings        = &settings;
    _shortcuts       = shortcuts;
    _shortcutManager = &shortcutManager;
    UpdateTheme(theme);

    const std::wstring title = LoadStringResource(nullptr, IDS_CMD_SHORTCUTS);

    if (owner && IsWindow(owner))
    {
        owner = GetAncestor(owner, GA_ROOT);
    }
    else
    {
        owner = nullptr;
    }

    const UINT dpi          = owner ? GetDpiForWindow(owner) : USER_DEFAULT_SCREEN_DPI;
    const int defaultWidth  = std::max(1, ThemedControls::ScaleDip(dpi, 820));
    const int defaultHeight = std::max(1, ThemedControls::ScaleDip(dpi, 520));

    const bool hasSavedPlacement = settings.windows.find(std::wstring(kShortcutsWindowId)) != settings.windows.end();

    CreateWindowExW(0,
                    kClassName,
                    title.c_str(),
                    WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
                    CW_USEDEFAULT,
                    CW_USEDEFAULT,
                    defaultWidth,
                    defaultHeight,
                    nullptr,
                    nullptr,
                    _hInstance,
                    this);

    if (! _hWnd)
    {
        return nullptr;
    }

    if (! hasSavedPlacement)
    {
        ResizeWindowToContent(_hWnd.get());
    }

    const int showCmd = hasSavedPlacement ? WindowPlacementPersistence::Restore(settings, kShortcutsWindowId, _hWnd.get()) : SW_SHOWNORMAL;
    ShowWindow(_hWnd.get(), showCmd);
    SetForegroundWindow(_hWnd.get());
    return _hWnd.get();
}

void ShortcutsWindow::UpdateTheme(const AppTheme& theme) noexcept
{
    _theme = theme;
    _backgroundBrush.reset(CreateSolidBrush(_theme.windowBackground));

    const COLORREF surface              = ThemedControls::GetControlSurfaceColor(_theme);
    _searchInputBackgroundColor         = ThemedControls::BlendColor(surface, _theme.windowBackground, _theme.dark ? 50 : 30, 255);
    _searchInputFocusedBackgroundColor  = ThemedControls::BlendColor(_searchInputBackgroundColor, _theme.menu.text, _theme.dark ? 20 : 16, 255);
    _searchInputDisabledBackgroundColor = ThemedControls::BlendColor(_theme.windowBackground, _searchInputBackgroundColor, _theme.dark ? 70 : 40, 255);

    _searchInputBrush.reset();
    _searchInputFocusedBrush.reset();
    _searchInputDisabledBrush.reset();
    if (! _theme.highContrast)
    {
        _searchInputBrush.reset(CreateSolidBrush(_searchInputBackgroundColor));
        _searchInputFocusedBrush.reset(CreateSolidBrush(_searchInputFocusedBackgroundColor));
        _searchInputDisabledBrush.reset(CreateSolidBrush(_searchInputDisabledBackgroundColor));
    }

    _searchFrameStyle.theme                        = &_theme;
    _searchFrameStyle.backdropBrush                = _backgroundBrush ? _backgroundBrush.get() : nullptr;
    _searchFrameStyle.inputBackgroundColor         = _searchInputBackgroundColor;
    _searchFrameStyle.inputFocusedBackgroundColor  = _searchInputFocusedBackgroundColor;
    _searchFrameStyle.inputDisabledBackgroundColor = _searchInputDisabledBackgroundColor;

    if (_hWnd)
    {
        ApplyTitleBarTheme(_hWnd.get(), _theme, GetActiveWindow() == _hWnd.get());
        ApplyListTheme();
        if (_searchFrame)
        {
            InvalidateRect(_searchFrame.get(), nullptr, TRUE);
        }
        if (_searchEdit)
        {
            InvalidateRect(_searchEdit, nullptr, TRUE);
        }
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void ShortcutsWindow::UpdateData(const Common::Settings::ShortcutsSettings& shortcuts, const ShortcutManager& shortcutManager) noexcept
{
    _shortcuts       = shortcuts;
    _shortcutManager = &shortcutManager;
    PopulateList();
    if (_hWnd)
    {
        const UINT dpi = GetDpiForWindow(_hWnd.get());
        AutoSizeColumnsToContent(dpi);
        ResizeWindowToContent(_hWnd.get());
    }
}

LRESULT CALLBACK ShortcutsWindow::WndProcThunk(HWND hWindow, UINT msg, WPARAM wp, LPARAM lp)
{
    ShortcutsWindow* self = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self    = reinterpret_cast<ShortcutsWindow*>(cs->lpCreateParams);
        SetWindowLongPtrW(hWindow, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        self->_hWnd.reset(hWindow);
        g_shortcutsWindow = self;
    }
    else
    {
        self = reinterpret_cast<ShortcutsWindow*>(GetWindowLongPtrW(hWindow, GWLP_USERDATA));
    }

    if (self)
    {
        return self->WndProc(hWindow, msg, wp, lp);
    }

    return DefWindowProcW(hWindow, msg, wp, lp);
}

LRESULT CALLBACK ShortcutsWindow::HeaderSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) noexcept
{
    auto* self = reinterpret_cast<ShortcutsWindow*>(dwRefData);
    if (! self || self->_theme.highContrast)
    {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    switch (msg)
    {
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: self->OnHeaderPaint(hwnd); return 0;
        case WM_NCDESTROY: return self->OnHeaderNcDestroy(hwnd, wParam, lParam, uIdSubclass);
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT ShortcutsWindow::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_PAINT: OnPaint(hwnd); return 0;
        case WM_ERASEBKGND: return 1;
        case WM_ACTIVATE: OnActivate(hwnd); return 0;
        case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, _theme, wp != FALSE); return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_GETMINMAXINFO: return OnGetMinMaxInfoMessage(hwnd, reinterpret_cast<MINMAXINFO*>(lp));
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_MEASUREITEM: return OnMeasureItem(reinterpret_cast<MEASUREITEMSTRUCT*>(lp));
        case WM_DRAWITEM: return OnDrawItem(reinterpret_cast<DRAWITEMSTRUCT*>(lp));
        case WM_DPICHANGED: return OnDpiChangedMessage(hwnd, static_cast<UINT>(HIWORD(wp)), reinterpret_cast<const RECT*>(lp));
        case WM_NOTIFY: return OnNotify(reinterpret_cast<const NMHDR*>(lp), lp);
        case WM_COMMAND: OnCommandMessage(LOWORD(wp), HIWORD(wp)); return 0;
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLOREDIT: return OnCtlColorEdit(reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CLOSE: DestroyWindow(hwnd); return 0;
        case WM_NCDESTROY: return OnNcDestroyMessage();
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ShortcutsWindow::OnHeaderNcDestroy(HWND header, WPARAM wParam, LPARAM lParam, UINT_PTR subclassId) noexcept
{
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
    RemoveWindowSubclass(header, HeaderSubclassProc, subclassId);
#pragma warning(pop)
    return DefSubclassProc(header, WM_NCDESTROY, wParam, lParam);
}

LRESULT ShortcutsWindow::OnDpiChangedMessage(HWND hwnd, UINT dpi, const RECT* suggested) noexcept
{
    if (suggested)
    {
        const int width  = static_cast<int>(std::max(0l, suggested->right - suggested->left));
        const int height = static_cast<int>(std::max(0l, suggested->bottom - suggested->top));
        SetWindowPos(hwnd, nullptr, suggested->left, suggested->top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    EnsureFonts(dpi);
    EnsureColumns(dpi);
    AutoSizeColumnsToContent(dpi);

    if (_searchEdit && ! _theme.highContrast)
    {
        ThemedControls::ApplyModernEditStyle(_searchEdit, _theme);
    }

    PopulateList();
    return 0;
}

LRESULT ShortcutsWindow::OnGetMinMaxInfoMessage(HWND hwnd, MINMAXINFO* info) noexcept
{
    if (! hwnd || ! info)
    {
        return 0;
    }

    const UINT dpi         = GetDpiForWindow(hwnd);
    const LONG minW        = static_cast<LONG>(ThemedControls::ScaleDip(dpi, 560));
    const LONG minH        = static_cast<LONG>(ThemedControls::ScaleDip(dpi, 420));
    info->ptMinTrackSize.x = std::max(info->ptMinTrackSize.x, minW);
    info->ptMinTrackSize.y = std::max(info->ptMinTrackSize.y, minH);

    static_cast<void>(WindowMaximizeBehavior::ApplyVerticalMaximize(hwnd, *info));
    return 0;
}

void ShortcutsWindow::EnsureFonts(UINT dpi) noexcept
{
    if (dpi == 0)
    {
        dpi = USER_DEFAULT_SCREEN_DPI;
    }

    if (_dpi == dpi && _uiFont)
    {
        return;
    }

    _dpi    = dpi;
    _uiFont = CreateMenuFontForDpi(dpi);

    HFONT fontToUse = _uiFont.get();
    if (! fontToUse)
    {
        fontToUse = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    if (_searchEdit)
    {
        SendMessageW(_searchEdit, WM_SETFONT, reinterpret_cast<WPARAM>(fontToUse), TRUE);
    }

    if (_list)
    {
        SendMessageW(_list, WM_SETFONT, reinterpret_cast<WPARAM>(fontToUse), TRUE);

        const HWND header = ListView_GetHeader(_list);
        if (header)
        {
            SendMessageW(header, WM_SETFONT, reinterpret_cast<WPARAM>(fontToUse), TRUE);
        }
    }
}

LRESULT ShortcutsWindow::OnNcDestroyMessage() noexcept
{
    if (_settings && _hWnd)
    {
        WindowPlacementPersistence::Save(*_settings, kShortcutsWindowId, _hWnd.get());

        const Common::Settings::Settings settingsToSave = SettingsSave::PrepareForSave(*_settings);
        const HRESULT saveHr                            = Common::Settings::SaveSettings(kSettingsAppId, settingsToSave);
        if (FAILED(saveHr))
        {
            const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kSettingsAppId);
            Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
        }
    }

    _hWnd.release();
    if (g_shortcutsWindow == this)
    {
        g_shortcutsWindow = nullptr;
    }
    delete this;
    return 0;
}

void ShortcutsWindow::Destroy() noexcept
{
    _backgroundBrush.reset();
    _searchInputBrush.reset();
    _searchInputFocusedBrush.reset();
    _searchInputDisabledBrush.reset();
    _imageList.reset();
    _rows.clear();
    _uiFont.reset();
    _searchFrame.reset();
    _searchEdit = nullptr;
    _searchQuery.clear();
    _list            = nullptr;
    _shortcutManager = nullptr;
    _hWnd.reset();
}

void ShortcutsWindow::OnCreate(HWND hwnd) noexcept
{
    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC  = ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icc);

    ApplyTitleBarTheme(hwnd, _theme, true);
    EnsureSearchControls(hwnd);
    EnsureListView(hwnd);

    const UINT dpi = GetDpiForWindow(hwnd);
    EnsureFonts(dpi);
    EnsureColumns(dpi);
    EnsureGroups();
    ApplyListTheme();
    PopulateList();
    AutoSizeColumnsToContent(dpi);
    ResizeWindowToContent(hwnd);

    if (_searchEdit)
    {
        SetFocus(_searchEdit);
    }
}

void ShortcutsWindow::OnDestroy() noexcept
{
}

void ShortcutsWindow::OnSize(UINT width, UINT height) noexcept
{
    if (! _list)
    {
        return;
    }

    const UINT dpi         = _hWnd ? GetDpiForWindow(_hWnd.get()) : USER_DEFAULT_SCREEN_DPI;
    const int padding      = ThemedControls::ScaleDip(dpi, 8);
    const int gapY         = ThemedControls::ScaleDip(dpi, 8);
    const int framePadding = std::max(1, ThemedControls::ScaleDip(dpi, 3));
    const int searchHeight = ThemedControls::ScaleDip(dpi, 34);

    int topY = 0;

    if (_searchEdit)
    {
        const int x = padding;
        const int y = padding;
        const int w = std::max(0, static_cast<int>(width) - 2 * padding);
        const int h = std::max(0, searchHeight);

        if (_searchFrame)
        {
            SetWindowPos(_searchFrame.get(), nullptr, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
        }

        SetWindowPos(_searchEdit,
                     nullptr,
                     x + framePadding,
                     y + framePadding,
                     std::max(1, w - 2 * framePadding),
                     std::max(1, h - 2 * framePadding),
                     SWP_NOZORDER | SWP_NOACTIVATE);

        topY = y + h + gapY;
    }

    MoveWindow(_list, 0, topY, static_cast<int>(width), std::max(0, static_cast<int>(height) - topY), TRUE);
}

void ShortcutsWindow::OnPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
    const HBRUSH bg           = _backgroundBrush ? _backgroundBrush.get() : static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH));
    FillRect(hdc.get(), &ps.rcPaint, bg);
}

void ShortcutsWindow::OnActivate(HWND /*hwnd*/) noexcept
{
    if (_list)
    {
        InvalidateRect(_list, nullptr, FALSE);
        const HWND header = ListView_GetHeader(_list);
        if (header)
        {
            InvalidateRect(header, nullptr, FALSE);
        }
    }
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ShortcutsWindow::OnHeaderPaint(HWND header) noexcept
{
    if (! header)
    {
        return;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(header, &ps);

    RECT client{};
    if (! GetClientRect(header, &client))
    {
        return;
    }

    const bool windowActive = _hWnd && GetActiveWindow() == _hWnd.get();
    const COLORREF bg       = BlendColor(_theme.windowBackground, _theme.menu.separator, 1, 12);
    COLORREF textColor      = windowActive ? _theme.menu.headerText : _theme.menu.headerTextDisabled;
    if (textColor == bg)
    {
        textColor = ChooseContrastingTextColor(bg);
    }

    wil::unique_hbrush bgBrush(CreateSolidBrush(bg));
    FillRect(hdc.get(), &ps.rcPaint, bgBrush.get());

    HFONT fontToUse = reinterpret_cast<HFONT>(SendMessageW(header, WM_GETFONT, 0, 0));
    if (! fontToUse)
    {
        fontToUse = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), fontToUse);

    const int dpi      = GetDeviceCaps(hdc.get(), LOGPIXELSX);
    const int paddingX = MulDiv(8, dpi, USER_DEFAULT_SCREEN_DPI);

    const COLORREF lineColor = _theme.menu.separator;
    wil::unique_hbrush lineBrush(CreateSolidBrush(lineColor));

    const int count = Header_GetItemCount(header);
    for (int i = 0; i < count; ++i)
    {
        RECT rc{};
        if (Header_GetItemRect(header, i, &rc) == FALSE)
        {
            continue;
        }

        if (! IntersectRect(&rc, &rc, &client))
        {
            continue;
        }

        wchar_t buf[128]{};
        HDITEMW item{};
        item.mask       = HDI_TEXT | HDI_FORMAT;
        item.pszText    = buf;
        item.cchTextMax = static_cast<int>(std::size(buf));
        if (Header_GetItem(header, i, &item) == FALSE)
        {
            continue;
        }

        RECT textRect  = rc;
        textRect.left  = std::min(textRect.right, textRect.left + paddingX);
        textRect.right = std::max(textRect.left, textRect.right - paddingX);

        UINT flags = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX;
        if ((item.fmt & HDF_RIGHT) != 0)
        {
            flags |= DT_RIGHT;
        }
        else if ((item.fmt & HDF_CENTER) != 0)
        {
            flags |= DT_CENTER;
        }
        else
        {
            flags |= DT_LEFT;
        }

        SetBkMode(hdc.get(), TRANSPARENT);
        SetTextColor(hdc.get(), textColor);
        DrawTextW(hdc.get(), buf, static_cast<int>(std::wcslen(buf)), &textRect, flags);

        RECT rightLine = rc;
        rightLine.left = std::max(rightLine.left, rightLine.right - 1);
        FillRect(hdc.get(), &rightLine, lineBrush.get());
    }

    RECT bottomLine = client;
    bottomLine.top  = std::max(bottomLine.top, bottomLine.bottom - 1);
    FillRect(hdc.get(), &bottomLine, lineBrush.get());
}

LRESULT ShortcutsWindow::OnMeasureItem(MEASUREITEMSTRUCT* mis) noexcept
{
    if (! mis || mis->CtlType != ODT_LISTVIEW || mis->CtlID != static_cast<UINT>(kListCtrlId))
    {
        return 0;
    }

    if (! _list)
    {
        return 0;
    }

    wil::unique_hdc_window hdc(GetDC(_list));
    if (! hdc)
    {
        return 1;
    }

    const HFONT font = reinterpret_cast<HFONT>(SendMessageW(_list, WM_GETFONT, 0, 0));
    if (font)
    {
        [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), font);
        mis->itemHeight               = static_cast<UINT>(std::max(1, GetRowHeightPx(hdc.get())));
        return 1;
    }

    mis->itemHeight = 36u;
    return 1;
}

LRESULT ShortcutsWindow::OnDrawItem(DRAWITEMSTRUCT* dis) noexcept
{
    if (! dis || dis->CtlType != ODT_LISTVIEW || dis->CtlID != static_cast<UINT>(kListCtrlId))
    {
        return 0;
    }

    if (! _list || ! dis->hDC)
    {
        return 1;
    }

    const int itemIndex = static_cast<int>(dis->itemID);
    if (itemIndex < 0)
    {
        return 1;
    }

    LVITEMW item{};
    item.mask  = LVIF_PARAM;
    item.iItem = itemIndex;
    if (ListView_GetItem(_list, &item) == FALSE)
    {
        return 1;
    }

    const size_t rowIndex = static_cast<size_t>(item.lParam);
    if (rowIndex >= _rows.size())
    {
        return 1;
    }

    const ShortcutRow& row = _rows[rowIndex];

    RECT rc = dis->rcItem;
    if (rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return 1;
    }

    const bool selected    = (dis->itemState & ODS_SELECTED) != 0;
    const bool focus       = (dis->itemState & ODS_FOCUS) != 0;
    const bool listFocused = _list && GetFocus() == _list;

    COLORREF bg = _theme.windowBackground;
    if (selected)
    {
        COLORREF selBg = _theme.menu.selectionBg;
        if (_theme.menu.rainbowMode && ! row.displayName.empty())
        {
            selBg = RainbowMenuSelectionColor(row.displayName, _theme.menu.darkBase);
        }

        if (listFocused || _theme.highContrast)
        {
            bg = selBg;
        }
        else
        {
            const int denom = _theme.menu.darkBase ? 2 : 3;
            bg              = BlendColor(_theme.windowBackground, selBg, 1, denom);
        }
    }
    else if (! _theme.highContrast && ((itemIndex % 2) == 1))
    {
        const COLORREF tint = _theme.menu.rainbowMode ? RainbowMenuSelectionColor(row.displayName, _theme.menu.darkBase) : _theme.menu.selectionBg;
        const int denom     = _theme.menu.darkBase ? 6 : 8;
        bg                  = BlendColor(_theme.windowBackground, tint, 1, denom);
    }

    wil::unique_hbrush bgBrush(CreateSolidBrush(bg));
    FillRect(dis->hDC, &rc, bgBrush.get());

    COLORREF textColor = selected ? ChooseContrastingTextColor(bg) : _theme.menu.text;
    if (textColor == bg)
    {
        textColor = ChooseContrastingTextColor(bg);
    }

    COLORREF descColor = textColor;
    if (! _theme.highContrast)
    {
        descColor = BlendColor(textColor, bg, 1, 3);
        if (descColor == bg)
        {
            descColor = textColor;
        }
    }

    constexpr int paddingX = 8;
    constexpr int paddingY = 3;
    constexpr int lineGap  = 1;

    const int commandColWidth = std::max(0, ListView_GetColumnWidth(_list, 0));
    RECT commandRect          = rc;
    commandRect.right         = std::min(rc.right, rc.left + commandColWidth);

    RECT keyRect = rc;
    keyRect.left = commandRect.right;

    int iconOffsetX = 0;
    if (row.conflict && _imageList)
    {
        constexpr int iconSize = 16;
        const int iconX        = commandRect.left + paddingX;
        const int iconY        = commandRect.top + std::max(0, (static_cast<int>(commandRect.bottom - commandRect.top) - iconSize) / 2);
        ImageList_Draw(_imageList.get(), 0, dis->hDC, iconX, iconY, ILD_NORMAL);
        iconOffsetX = iconSize + 6;
    }

    RECT textRect   = commandRect;
    textRect.left   = std::min(textRect.right, textRect.left + paddingX + iconOffsetX);
    textRect.right  = std::max(textRect.left, textRect.right - paddingX);
    textRect.top    = std::min(textRect.bottom, textRect.top + paddingY);
    textRect.bottom = std::max(textRect.top, textRect.bottom - paddingY);

    TEXTMETRICW tm{};
    GetTextMetricsW(dis->hDC, &tm);
    const int lineHeight = std::max(1, static_cast<int>(tm.tmHeight + tm.tmExternalLeading));

    RECT nameRect   = textRect;
    nameRect.bottom = std::min(textRect.bottom, nameRect.top + lineHeight);

    RECT descRect = textRect;
    descRect.top  = std::min(textRect.bottom, nameRect.bottom + lineGap);

    SetBkMode(dis->hDC, TRANSPARENT);

    const std::wstring_view query        = _searchQuery;
    const std::wstring_view trimmedQuery = TrimWhitespace(query);

    COLORREF highlightBg        = bg;
    COLORREF highlightTextColor = textColor;
    if (! trimmedQuery.empty())
    {
        if (_theme.highContrast)
        {
            highlightBg = GetSysColor(COLOR_HIGHLIGHT);
        }
        else
        {
            const int denom = _theme.menu.darkBase ? 2 : 3;
            highlightBg     = BlendColor(bg, _theme.menu.selectionBg, 1, denom);
            if (highlightBg == bg)
            {
                highlightBg = BlendColor(bg, textColor, 1, _theme.menu.darkBase ? 4 : 6);
            }
        }
        highlightTextColor = ChooseContrastingTextColor(highlightBg);
    }

    HBRUSH highlightBrush = nullptr;
    wil::unique_hbrush ownedHighlightBrush;
    if (! trimmedQuery.empty() && ! _theme.highContrast)
    {
        ownedHighlightBrush.reset(CreateSolidBrush(highlightBg));
        highlightBrush = ownedHighlightBrush.get();
    }
    else if (! trimmedQuery.empty() && _theme.highContrast)
    {
        highlightBrush = GetSysColorBrush(COLOR_HIGHLIGHT);
    }

    DrawTextWithHighlights(
        dis->hDC, row.displayName, nameRect, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS, query, textColor, highlightTextColor, highlightBrush);

    if (! row.description.empty())
    {
        DrawTextWithHighlights(
            dis->hDC, row.description, descRect, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS, query, descColor, highlightTextColor, highlightBrush);
    }

    RECT keyTextRect  = keyRect;
    keyTextRect.left  = std::min(keyTextRect.right, keyTextRect.left + paddingX);
    keyTextRect.right = std::max(keyTextRect.left, keyTextRect.right - paddingX);

    DrawTextWithHighlights(dis->hDC,
                           row.keyText,
                           keyTextRect,
                           DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS,
                           query,
                           textColor,
                           highlightTextColor,
                           highlightBrush);

    if (focus)
    {
        DrawFocusRect(dis->hDC, &rc);
    }

    return 1;
}

void ShortcutsWindow::OnCommandMessage(UINT controlId, UINT notifyCode) noexcept
{
    if (controlId != static_cast<UINT>(kSearchEditId))
    {
        return;
    }

    if (notifyCode == EN_CHANGE)
    {
        OnSearchChanged();
    }
}

LRESULT ShortcutsWindow::OnCtlColorEdit(HDC hdc, HWND control) noexcept
{
    if (! hdc || ! control)
    {
        return 0;
    }

    if (_theme.highContrast || control != _searchEdit)
    {
        return 0;
    }

    const bool enabled = IsWindowEnabled(control) != FALSE;
    const bool focused = GetFocus() == control;

    COLORREF bg  = _searchInputBackgroundColor;
    HBRUSH brush = _searchInputBrush.get();
    if (! enabled)
    {
        bg    = _searchInputDisabledBackgroundColor;
        brush = _searchInputDisabledBrush.get();
    }
    else if (focused)
    {
        bg    = _searchInputFocusedBackgroundColor;
        brush = _searchInputFocusedBrush.get();
    }

    if (! brush)
    {
        return 0;
    }

    COLORREF textColor = _theme.menu.text;
    if (textColor == bg)
    {
        textColor = ChooseContrastingTextColor(bg);
    }

    SetBkColor(hdc, bg);
    SetTextColor(hdc, textColor);
    return reinterpret_cast<LRESULT>(brush);
}

void ShortcutsWindow::OnSearchChanged() noexcept
{
    if (! _searchEdit)
    {
        return;
    }

    const std::wstring text         = GetWindowTextString(_searchEdit);
    const std::wstring_view trimmed = TrimWhitespace(text);

    std::wstring newQuery(trimmed);
    if (newQuery == _searchQuery)
    {
        return;
    }

    _searchQuery = std::move(newQuery);
    PopulateList();
}

LRESULT ShortcutsWindow::OnNotify(const NMHDR* header, LPARAM lp) noexcept
{
    if (! header || ! _list || header->hwndFrom != _list)
    {
        return 0;
    }

    if (header->code == NM_CUSTOMDRAW)
    {
        return OnCustomDraw(reinterpret_cast<NMLVCUSTOMDRAW*>(lp));
    }

    if (header->code == LVN_GETINFOTIPW)
    {
        return OnGetInfoTip(reinterpret_cast<NMLVGETINFOTIPW*>(lp));
    }

    return 0;
}

LRESULT ShortcutsWindow::OnCustomDraw(NMLVCUSTOMDRAW* cd) noexcept
{
    if (! cd)
    {
        return CDRF_DODEFAULT;
    }

    if (cd->nmcd.dwDrawStage == CDDS_PREPAINT)
    {
        return CDRF_NOTIFYITEMDRAW;
    }

    if (cd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT || cd->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM))
    {
        if (cd->dwItemType == LVCDI_GROUP)
        {
            cd->clrTextBk           = _theme.windowBackground;
            const bool windowActive = _hWnd && GetActiveWindow() == _hWnd.get();
            COLORREF text           = windowActive ? _theme.menu.headerText : _theme.menu.headerTextDisabled;
            if (text == _theme.windowBackground)
            {
                text = ChooseContrastingTextColor(_theme.windowBackground);
            }
            cd->clrText = text;
            return CDRF_NEWFONT;
        }
    }

    return CDRF_DODEFAULT;
}

LRESULT ShortcutsWindow::OnGetInfoTip(NMLVGETINFOTIPW* tip) noexcept
{
    if (! tip || tip->iItem < 0 || ! tip->pszText || tip->cchTextMax <= 0)
    {
        return 0;
    }

    LVITEMW item{};
    item.mask  = LVIF_PARAM;
    item.iItem = tip->iItem;
    if (ListView_GetItem(_list, &item) == FALSE)
    {
        return 0;
    }

    const size_t rowIndex = static_cast<size_t>(item.lParam);
    if (rowIndex >= _rows.size())
    {
        return 0;
    }

    const ShortcutRow& row = _rows[rowIndex];
    if (! row.conflict || row.conflictWith.empty())
    {
        return 0;
    }

    const std::wstring chordText = FormatChordText(row.binding.vk, row.binding.modifiers);
    const std::wstring text      = FormatStringResource(nullptr, IDS_FMT_SHORTCUT_CONFLICT, row.conflictWith, chordText);

    wcsncpy_s(tip->pszText, static_cast<size_t>(tip->cchTextMax), text.c_str(), _TRUNCATE);
    return 0;
}

void ShortcutsWindow::EnsureSearchControls(HWND hwnd) noexcept
{
    if (_searchEdit)
    {
        return;
    }

    if (! hwnd)
    {
        return;
    }

    const DWORD exStyle = _theme.highContrast ? WS_EX_CLIENTEDGE : 0;
    _searchEdit         = CreateWindowExW(exStyle,
                                  L"Edit",
                                  _searchQuery.c_str(),
                                  WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
                                  0,
                                  0,
                                  10,
                                  10,
                                  hwnd,
                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(kSearchEditId)),
                                  _hInstance,
                                  nullptr);

    if (! _searchEdit)
    {
        return;
    }

    const std::wstring cue = LoadStringResource(nullptr, IDS_SHORTCUTS_SEARCH_CUE);
    if (! cue.empty())
    {
        SendMessageW(_searchEdit, EM_SETCUEBANNER, TRUE, reinterpret_cast<LPARAM>(cue.c_str()));
    }

    if (_theme.highContrast)
    {
        return;
    }

    ThemedControls::ApplyModernEditStyle(_searchEdit, _theme);

    _searchFrame.reset(CreateWindowExW(0, L"Static", L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS, 0, 0, 10, 10, hwnd, nullptr, _hInstance, nullptr));

    if (! _searchFrame)
    {
        return;
    }

    SetWindowPos(_searchFrame.get(), _searchEdit, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    ThemedInputFrames::InstallFrame(_searchFrame.get(), _searchEdit, &_searchFrameStyle);
}

void ShortcutsWindow::EnsureListView(HWND hwnd) noexcept
{
    if (_list)
    {
        return;
    }

    _list = CreateWindowExW(0,
                            WC_LISTVIEWW,
                            L"",
                            WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_OWNERDRAWFIXED | LVS_SHOWSELALWAYS,
                            0,
                            0,
                            0,
                            0,
                            hwnd,
                            reinterpret_cast<HMENU>(static_cast<INT_PTR>(kListCtrlId)),
                            _hInstance,
                            nullptr);

    if (! _list)
    {
        return;
    }

    ListView_SetExtendedListViewStyle(_list, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP | LVS_EX_INFOTIP);

    _imageList.reset(ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 1, 1));
    if (_imageList)
    {
        wil::unique_hicon warnIcon(static_cast<HICON>(LoadImageW(nullptr, IDI_WARNING, IMAGE_ICON, 16, 16, 0)));
        if (warnIcon)
        {
            ImageList_AddIcon(_imageList.get(), warnIcon.get());
        }
    }

    ListView_SetImageList(_list, _imageList.get(), LVSIL_SMALL);
    ListView_EnableGroupView(_list, TRUE);
}

void ShortcutsWindow::EnsureColumns(UINT dpi) noexcept
{
    if (! _list)
    {
        return;
    }

    ListView_DeleteAllItems(_list);

    while (ListView_DeleteColumn(_list, 0))
    {
    }

    const auto add = [&](int index, UINT textId, int widthDip, int fmt = LVCFMT_LEFT) noexcept
    {
        const std::wstring text = LoadStringResource(nullptr, textId);
        LVCOLUMNW col{};
        col.mask    = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT;
        col.pszText = const_cast<wchar_t*>(text.c_str());
        col.cx      = MulDiv(widthDip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
        col.fmt     = fmt;
        ListView_InsertColumn(_list, index, &col);
    };

    add(0, IDS_SHORTCUTS_COL_COMMAND, 520);
    add(1, IDS_SHORTCUTS_COL_KEY, 260, LVCFMT_RIGHT);
}

void ShortcutsWindow::EnsureGroups() noexcept
{
    if (! _list)
    {
        return;
    }

    ListView_RemoveAllGroups(_list);

    const auto addGroup = [&](int groupId, UINT titleId) noexcept
    {
        const std::wstring title = LoadStringResource(nullptr, titleId);
        LVGROUP group{};
        group.cbSize    = sizeof(group);
        group.mask      = LVGF_GROUPID | LVGF_HEADER;
        group.iGroupId  = groupId;
        group.pszHeader = const_cast<wchar_t*>(title.c_str());
        ListView_InsertGroup(_list, -1, &group);
    };

    addGroup(kGroupFunctionBar, IDS_SHORTCUTS_GROUP_FUNCTION_BAR);
    addGroup(kGroupFolderView, IDS_SHORTCUTS_GROUP_FOLDER_VIEW);
}

void ShortcutsWindow::ApplyListTheme() noexcept
{
    if (! _list)
    {
        return;
    }

    ListView_SetBkColor(_list, _theme.windowBackground);
    ListView_SetTextBkColor(_list, _theme.windowBackground);
    ListView_SetTextColor(_list, _theme.menu.text);

    const bool darkBackground = ChooseContrastingTextColor(_theme.windowBackground) == RGB(255, 255, 255);
    const wchar_t* listTheme  = (_theme.highContrast ? L"" : (darkBackground ? L"DarkMode_Explorer" : L"Explorer"));
    SetWindowTheme(_list, listTheme, nullptr);

    const HWND header = ListView_GetHeader(_list);
    if (header)
    {
        SetWindowTheme(header, listTheme, nullptr);
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(header, HeaderSubclassProc, kListHeaderSubclassId, reinterpret_cast<DWORD_PTR>(this));
#pragma warning(pop)
        InvalidateRect(header, nullptr, TRUE);
    }

    const HWND tooltips = ListView_GetToolTips(_list);
    if (tooltips)
    {
        SetWindowTheme(tooltips, listTheme, nullptr);
    }

    if (_searchEdit && ! _theme.highContrast)
    {
        SetWindowTheme(_searchEdit, listTheme, nullptr);
    }
}

void ShortcutsWindow::PopulateList() noexcept
{
    if (! _list || ! _shortcutManager)
    {
        return;
    }

    ListView_DeleteAllItems(_list);
    _rows.clear();

    auto addScope = [&](const std::vector<Common::Settings::ShortcutBinding>& bindings, const std::vector<uint32_t>& conflicts, int groupId) noexcept
    {
        std::unordered_map<uint32_t, std::vector<size_t>> chordToRows;

        for (const auto& binding : bindings)
        {
            ShortcutRow row;
            row.binding     = binding;
            row.displayName = GetCommandDisplayName(binding.commandId);
            row.description = GetCommandDescription(binding.commandId);
            row.keyText     = FormatChordText(binding.vk, binding.modifiers);
            row.chordKey    = ShortcutManager::MakeChordKey(binding.vk, binding.modifiers);
            row.groupId     = groupId;
            row.conflict    = IsConflictChord(row.chordKey, conflicts);

            const size_t rowIndex = _rows.size();
            _rows.push_back(std::move(row));
            chordToRows[_rows[rowIndex].chordKey].push_back(rowIndex);
        }

        for (const auto& [_, indices] : chordToRows)
        {
            if (indices.size() <= 1)
            {
                continue;
            }

            for (size_t i = 0; i < indices.size(); ++i)
            {
                const size_t idx   = indices[i];
                const size_t other = indices[(i + 1) % indices.size()];
                if (idx < _rows.size() && other < _rows.size())
                {
                    _rows[idx].conflictWith = _rows[other].displayName;
                }
            }
        }
    };

    addScope(_shortcuts.functionBar, _shortcutManager->GetFunctionBarConflicts(), kGroupFunctionBar);
    addScope(_shortcuts.folderView, _shortcutManager->GetFolderViewConflicts(), kGroupFolderView);

    const std::wstring_view query = TrimWhitespace(_searchQuery);
    const bool filterEnabled      = ! query.empty();

    int listIndex = 0;
    for (size_t rowIndex = 0; rowIndex < _rows.size(); ++rowIndex)
    {
        ShortcutRow& row = _rows[rowIndex];
        if (filterEnabled)
        {
            const bool matches = ContainsNoCase(row.displayName, query) || ContainsNoCase(row.description, query) || ContainsNoCase(row.keyText, query);
            if (! matches)
            {
                continue;
            }
        }

        LVITEMW item{};
        item.mask     = LVIF_TEXT | LVIF_PARAM | LVIF_IMAGE | LVIF_GROUPID;
        item.iItem    = listIndex;
        item.iSubItem = 0;
        item.pszText  = const_cast<wchar_t*>(row.displayName.c_str());
        item.lParam   = static_cast<LPARAM>(rowIndex);
        item.iImage   = row.conflict ? 0 : I_IMAGENONE;
        item.iGroupId = row.groupId;

        const int inserted = ListView_InsertItem(_list, &item);
        if (inserted < 0)
        {
            continue;
        }

        ListView_SetItemText(_list, inserted, 1, const_cast<wchar_t*>(row.keyText.c_str()));
        ++listIndex;
    }
}

void ShortcutsWindow::ResizeWindowToContent(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (! _list)
    {
        return;
    }

    HMONITOR mon = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (! GetMonitorInfoW(mon, &mi))
    {
        return;
    }

    const RECT& rcWork = mi.rcWork;
    const int workW    = static_cast<int>(std::max(0l, rcWork.right - rcWork.left));
    const int workH    = static_cast<int>(std::max(0l, rcWork.bottom - rcWork.top));
    if (workW <= 0 || workH <= 0)
    {
        return;
    }

    RECT windowRect{};
    RECT clientRect{};
    if (! GetWindowRect(hwnd, &windowRect) || ! GetClientRect(hwnd, &clientRect))
    {
        return;
    }

    const int nonClientW = static_cast<int>(std::max(0l, (windowRect.right - windowRect.left) - (clientRect.right - clientRect.left)));

    const UINT dpi    = GetDpiForWindow(hwnd);
    const int scrollW = GetSystemMetricsForDpi(SM_CXVSCROLL, dpi);

    const int listItems   = ListView_GetItemCount(_list);
    const int perPage     = ListView_GetCountPerPage(_list);
    const bool hasVScroll = perPage > 0 && listItems > perPage;

    int desiredListClientW = std::max(0, ListView_GetColumnWidth(_list, 0) + ListView_GetColumnWidth(_list, 1));
    if (hasVScroll)
    {
        desiredListClientW += scrollW;
    }

    const int minWindowW = MulDiv(640, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
    const int maxWindowW = std::max(minWindowW, MulDiv(1200, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));

    int desiredWindowW = desiredListClientW + nonClientW;
    desiredWindowW     = std::clamp(desiredWindowW, minWindowW, std::min(workW, maxWindowW));

    const int desiredWindowH = workH;
    SetWindowPos(hwnd, nullptr, rcWork.left, rcWork.top, desiredWindowW, desiredWindowH, SWP_NOZORDER | SWP_NOACTIVATE);
}

int ShortcutsWindow::GetRowHeightPx(HDC hdc) const noexcept
{
    if (! hdc)
    {
        return 36;
    }

    TEXTMETRICW tm{};
    if (! GetTextMetricsW(hdc, &tm))
    {
        return 36;
    }

    const int lineHeight   = std::max(1, static_cast<int>(tm.tmHeight + tm.tmExternalLeading));
    constexpr int paddingY = 3;
    constexpr int lineGap  = 1;
    return (paddingY * 2) + (lineHeight * 2) + lineGap;
}

void ShortcutsWindow::AutoSizeColumnsToContent(UINT dpi) noexcept
{
    if (! _list)
    {
        return;
    }

    wil::unique_hdc_window hdc(GetDC(_list));
    if (! hdc)
    {
        return;
    }

    const HFONT font = reinterpret_cast<HFONT>(SendMessageW(_list, WM_GETFONT, 0, 0));
    if (font)
    {
        [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), font);
    }

    int maxCommand = 0;
    int maxKey     = 0;

    auto measure = [&](const std::wstring& text, int& outMax) noexcept
    {
        if (text.empty())
        {
            return;
        }

        SIZE s{};
        GetTextExtentPoint32W(hdc.get(), text.c_str(), static_cast<int>(text.size()), &s);
        outMax = std::max(outMax, static_cast<int>(s.cx));
    };

    measure(LoadStringResource(nullptr, IDS_SHORTCUTS_COL_COMMAND), maxCommand);
    measure(LoadStringResource(nullptr, IDS_SHORTCUTS_COL_KEY), maxKey);

    for (const auto& row : _rows)
    {
        measure(row.displayName, maxCommand);
        measure(row.description, maxCommand);
        measure(row.keyText, maxKey);
    }

    const int paddingX      = MulDiv(16, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
    constexpr int iconSpace = 22;
    int desiredCommandWidth = maxCommand + paddingX + iconSpace;
    int desiredKeyWidth     = maxKey + paddingX;

    RECT client{};
    if (! GetClientRect(_list, &client))
    {
        return;
    }

    const int listItems   = ListView_GetItemCount(_list);
    const int perPage     = ListView_GetCountPerPage(_list);
    const bool hasVScroll = perPage > 0 && listItems > perPage;

    int available = static_cast<int>(std::max(0l, client.right - client.left));
    if (hasVScroll)
    {
        available = std::max(0, available - GetSystemMetricsForDpi(SM_CXVSCROLL, dpi));
    }

    const int minKeyWidth     = MulDiv(160, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
    const int minCommandWidth = MulDiv(260, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);

    desiredKeyWidth     = std::max(desiredKeyWidth, minKeyWidth);
    desiredCommandWidth = std::max(desiredCommandWidth, minCommandWidth);

    int keyWidth     = desiredKeyWidth;
    int commandWidth = desiredCommandWidth;

    if (available > 0)
    {
        if ((commandWidth + keyWidth) > available)
        {
            keyWidth     = std::min(keyWidth, std::max(minKeyWidth, available / 2));
            commandWidth = std::max(minCommandWidth, available - keyWidth);
            if ((commandWidth + keyWidth) > available)
            {
                keyWidth     = std::max(minKeyWidth, available - minCommandWidth);
                commandWidth = std::max(minCommandWidth, available - keyWidth);
            }
        }
        else
        {
            commandWidth = std::max(minCommandWidth, available - keyWidth);
        }
    }

    ListView_SetColumnWidth(_list, 0, commandWidth);
    ListView_SetColumnWidth(_list, 1, keyWidth);
}

} // namespace

void ShowShortcutsWindow(HWND owner,
                         Common::Settings::Settings& settings,
                         const Common::Settings::ShortcutsSettings& shortcuts,
                         const ShortcutManager& shortcutManager,
                         const AppTheme& theme) noexcept
{
    if (g_shortcutsWindow && g_shortcutsWindow->GetHwnd())
    {
        g_shortcutsWindow->UpdateData(shortcuts, shortcutManager);
        g_shortcutsWindow->UpdateTheme(theme);

        const HWND hwnd = g_shortcutsWindow->GetHwnd();
        if (IsIconic(hwnd))
        {
            ShowWindow(hwnd, SW_RESTORE);
        }
        else
        {
            ShowWindow(hwnd, SW_SHOW);
        }
        SetForegroundWindow(hwnd);
        return;
    }

    auto window = std::make_unique<ShortcutsWindow>();
    if (window->Create(owner, settings, shortcuts, shortcutManager, theme))
    {
        // Window successfully created - it will self-delete in WM_NCDESTROY
        static_cast<void>(window.release());
    }
}

void UpdateShortcutsWindowTheme(const AppTheme& theme) noexcept
{
    if (! g_shortcutsWindow)
    {
        return;
    }

    g_shortcutsWindow->UpdateTheme(theme);
}

void UpdateShortcutsWindowData(const Common::Settings::ShortcutsSettings& shortcuts, const ShortcutManager& shortcutManager) noexcept
{
    if (! g_shortcutsWindow)
    {
        return;
    }

    g_shortcutsWindow->UpdateData(shortcuts, shortcutManager);
}

HWND GetShortcutsWindowHandle() noexcept
{
    if (! g_shortcutsWindow)
    {
        return nullptr;
    }

    const HWND hwnd = g_shortcutsWindow->GetHwnd();
    if (! hwnd || ! IsWindow(hwnd))
    {
        return nullptr;
    }

    return hwnd;
}
