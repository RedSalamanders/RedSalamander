#include "Framework.h"

#include "CommandPaletteWindow.h"

#include "CommandRuntimeState.h"
#include "CommandVisuals.h"
#include "DxUiThemePalette.h"
#include "Helpers.h"
#include "ShortcutCommandCatalog.h"
#include "resource.h"

#include <algorithm>
#include <cwctype>

[[nodiscard]] bool DispatchShortcutCommandFromWindow(HWND ownerWindow, std::wstring_view commandId) noexcept;

namespace
{
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::GridSortSpec;
using RedSalamander::DxUi::IDxGridDelegate;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::WindowHost;

constexpr wchar_t kClassName[] = L"RedSalamander.CommandPaletteWindow";

[[nodiscard]] bool ContainsNoCase(std::wstring_view text, std::wstring_view query) noexcept
{
    return query.empty() || std::search(text.begin(), text.end(), query.begin(), query.end(), [](wchar_t left, wchar_t right) noexcept {
        return std::towupper(static_cast<wint_t>(left)) == std::towupper(static_cast<wint_t>(right));
    }) != text.end();
}

[[nodiscard]] bool StartsWithNoCase(std::wstring_view text, std::wstring_view query) noexcept
{
    return query.size() <= text.size() &&
           CompareStringOrdinal(text.data(), static_cast<int>(query.size()), query.data(), static_cast<int>(query.size()), TRUE) == CSTR_EQUAL;
}

struct PaletteRow final
{
    ShortcutCommandCatalogEntry command;
    CommandRuntimeState runtimeState{};
    std::wstring commandText;
    std::wstring shortcutText;
    std::wstring iconText;
    uint32_t rank = 0u;
};

class PaletteGridModel final : public IDxGridModel
{
public:
    PaletteGridModel()
    {
        _columns = {
            {L"command", LoadStringResource(nullptr, IDS_SHORTCUTS_COL_COMMAND), 530.0f, 260.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
            {L"shortcut", LoadStringResource(nullptr, IDS_SHORTCUTS_COL_KEY), 250.0f, 120.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
        };
    }

    void SetRows(std::vector<PaletteRow> rows)
    {
        _rows = std::move(rows);
    }
    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rows.size();
    }
    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }
    [[nodiscard]] GridColumnDesc GetColumn(size_t index) const override
    {
        return _columns.at(index);
    }
    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& cell) const override
    {
        cell = {};
        if (rowIndex >= _rows.size())
            return;
        const PaletteRow& row = _rows[rowIndex];
        if (columnIndex == 0u)
        {
            cell.kind        = RedSalamander::DxUi::GridCellKind::IconText;
            cell.iconText    = row.iconText;
            cell.text        = row.commandText;
            cell.multiline   = true;
            cell.tooltipText = row.command.searchText;
            cell.enabled     = row.runtimeState.enabled;
        }
        else
        {
            cell.text          = row.shortcutText;
            cell.textAlignment = DWRITE_TEXT_ALIGNMENT_TRAILING;
            cell.tooltipText   = row.shortcutText;
            cell.enabled       = row.runtimeState.enabled;
        }
    }
    [[nodiscard]] RedSalamander::DxUi::GridRowStyle GetRowStyle(size_t index) const override
    {
        return index < _rows.size() ? RedSalamander::DxUi::GridRowStyle{.rainbowSeed = _rows[index].command.commandId} : RedSalamander::DxUi::GridRowStyle{};
    }
    [[nodiscard]] uint64_t GetStableRowId(size_t index) const noexcept override
    {
        return index + 1u;
    }
    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t id) const noexcept override
    {
        return id != 0u && id <= _rows.size() ? std::optional<size_t>(static_cast<size_t>(id - 1u)) : std::nullopt;
    }
    [[nodiscard]] std::wstring GetRowAccessibleName(size_t index) const
    {
        return index < _rows.size() ? _rows[index].commandText : std::wstring{};
    }
    [[nodiscard]] const PaletteRow* GetRow(size_t index) const noexcept
    {
        return index < _rows.size() ? &_rows[index] : nullptr;
    }
    void SortRows(const GridSortSpec&)
    {
    }

private:
    std::vector<GridColumnDesc> _columns;
    std::vector<PaletteRow> _rows;
};

class CommandPaletteWindow final : public IDxGridDelegate
{
public:
    CommandPaletteWindow()                                       = default;
    CommandPaletteWindow(const CommandPaletteWindow&)            = delete;
    CommandPaletteWindow(CommandPaletteWindow&&)                 = delete;
    CommandPaletteWindow& operator=(const CommandPaletteWindow&) = delete;
    CommandPaletteWindow& operator=(CommandPaletteWindow&&)      = delete;
    [[nodiscard]] HWND Create(HWND owner, const Common::Settings::ShortcutsSettings& shortcuts, bool terminalContext, const AppTheme& theme) noexcept;
    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _window.get();
    }
    void UpdateTheme(const AppTheme& theme) noexcept;
    void OnGridSelectionChanged(Grid&) override
    {
    }
    void OnGridSelectionChanged() override
    {
    }
    void OnGridSortRequested(const GridSortSpec&) override
    {
    }
    void OnGridRowActivated(Grid&, size_t index) override
    {
        Activate(index);
    }
    void OnGridRowActivated(size_t index) override
    {
        Activate(index);
    }
#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugSnapshot(CommandPaletteDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugSetSearch(std::wstring_view text) noexcept;
#endif

private:
    static ATOM RegisterClass(HINSTANCE instance) noexcept;
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    LRESULT HandleMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    void BuildUi();
    void RebuildRows();
    void Layout() noexcept;
    void ApplyTheme() noexcept;
    void Activate(size_t index) noexcept;

    wil::unique_hwnd _window;
    HWND _owner       = nullptr;
    HWND _returnFocus = nullptr;
    AppTheme _theme{};
    bool _terminalContext = false;
    std::wstring _search;
    std::vector<ShortcutCommandCatalogEntry> _catalog;
    WindowHost _host;
    std::unique_ptr<Panel> _rootStorage;
    Panel* _root            = nullptr;
    TextField* _searchField = nullptr;
    Grid* _grid             = nullptr;
    Label* _emptyLabel      = nullptr;
    std::unique_ptr<PaletteGridModel> _modelStorage;
    PaletteGridModel* _model = nullptr;
};

std::unique_ptr<CommandPaletteWindow> g_palette;

ATOM CommandPaletteWindow::RegisterClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom != 0u)
        return atom;
    WNDCLASSEXW windowClass{};
    windowClass.cbSize        = sizeof(windowClass);
    windowClass.lpfnWndProc   = &CommandPaletteWindow::WindowProc;
    windowClass.hInstance     = instance;
    windowClass.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    windowClass.hIcon         = LoadIconW(instance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
    windowClass.hbrBackground = nullptr;
    windowClass.lpszClassName = kClassName;
    atom                      = RegisterClassExW(&windowClass);
    return atom;
}

HWND CommandPaletteWindow::Create(HWND owner, const Common::Settings::ShortcutsSettings& shortcuts, bool terminalContext, const AppTheme& theme) noexcept
{
    const auto openStartedAt = std::chrono::steady_clock::now();
    _owner                   = owner;
    _returnFocus             = GetFocus();
    _theme                   = theme;
    _terminalContext         = terminalContext;
    _catalog                 = BuildShortcutCommandCatalog(shortcuts, terminalContext ? ShortcutCommandContext::Terminal : ShortcutCommandContext::Folder);
    const HINSTANCE instance = GetModuleHandleW(nullptr);
    if (! RegisterClass(instance))
        return nullptr;

    const UINT dpi   = owner ? GetDpiForWindow(owner) : GetDpiForSystem();
    const int width  = MulDiv(820, static_cast<int>(std::max(dpi, 96u)), 96);
    const int height = MulDiv(560, static_cast<int>(std::max(dpi, 96u)), 96);
    RECT ownerRect{};
    GetWindowRect(owner, &ownerRect);
    const int x     = static_cast<int>(ownerRect.left + std::max<LONG>(0L, (ownerRect.right - ownerRect.left - width) / 2));
    const int y     = static_cast<int>(ownerRect.top + std::max<LONG>(0L, (ownerRect.bottom - ownerRect.top - height) / 3));
    const HWND hwnd = CreateWindowExW(WS_EX_TOOLWINDOW,
                                      kClassName,
                                      LoadStringResource(nullptr, IDS_CMD_COMMAND_PALETTE).c_str(),
                                      WS_POPUP | WS_THICKFRAME | WS_CLIPCHILDREN,
                                      x,
                                      y,
                                      width,
                                      height,
                                      owner,
                                      nullptr,
                                      instance,
                                      this);
    if (! hwnd)
        return nullptr;
    if (! _window)
        _window.reset(hwnd);
    ShowWindow(hwnd, SW_SHOWNORMAL);
    SetForegroundWindow(hwnd);
    Debug::Perf::EmitDurationUs(L"app.command_palette.open_to_visible_us", Debug::Perf::ElapsedUs(openStartedAt), _catalog.size(), 0u);
    return hwnd;
}

LRESULT CALLBACK CommandPaletteWindow::WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message == WM_NCCREATE)
    {
        const auto* create = reinterpret_cast<CREATESTRUCTW*>(lParam);
        auto* self         = static_cast<CommandPaletteWindow*>(create ? create->lpCreateParams : nullptr);
        if (! self)
            return FALSE;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        if (! self->_window)
            self->_window.reset(hwnd);
    }
    auto* self = reinterpret_cast<CommandPaletteWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return self ? self->HandleMessage(hwnd, message, wParam, lParam) : DefWindowProcW(hwnd, message, wParam, lParam);
}

LRESULT CommandPaletteWindow::HandleMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    // Enter and Escape are palette-level commands. Handle them before the
    // focused DxUi text field can consume the key as ordinary edit input.
    if (message == WM_KEYDOWN && wParam == VK_ESCAPE)
    {
        _window.reset();
        if (_returnFocus != nullptr && IsWindow(_returnFocus) != FALSE)
        {
            SetFocus(_returnFocus);
        }
        return 0;
    }
    if (message == WM_KEYDOWN && wParam == VK_RETURN && _grid && _model)
    {
        const auto selected = _grid->GetPrimarySelectedRow();
        if (selected.has_value())
            Activate(selected.value());
        return 0;
    }

    bool handled             = false;
    const LRESULT hostResult = message == WM_CREATE ? 0 : _host.HandleMessage(hwnd, message, wParam, lParam, handled);
    if (handled)
    {
        if (message == WM_SIZE || message == WM_DPICHANGED)
        {
            Layout();
        }
        else if (message == WM_ACTIVATE && LOWORD(wParam) == WA_INACTIVE && _window)
        {
            _window.reset();
            return 0;
        }
        return hostResult;
    }
    switch (message)
    {
        case WM_CREATE:
            if (! _host.Attach(hwnd))
                return -1;
            BuildUi();
            ApplyTheme();
            RebuildRows();
            Layout();
            _host.SetFocusControl(_searchField);
            return 0;
        case WM_SIZE:
        case WM_DPICHANGED: Layout(); return 0;
        case WM_ACTIVATE:
            if (LOWORD(wParam) == WA_INACTIVE && _window)
            {
                _window.reset();
                return 0;
            }
            break;
        case WM_ERASEBKGND: return 1;
        case WM_CLOSE:
            _window.reset();
            if (_returnFocus != nullptr && IsWindow(_returnFocus) != FALSE)
            {
                SetFocus(_returnFocus);
            }
            return 0;
        case WM_KEYDOWN:
            if (wParam == VK_DOWN && _grid && _model && _model->GetRowCount() != 0u)
            {
                _host.SetFocusControl(_grid);
                static_cast<void>(_grid->RequestSelectRow(0u, 0u));
                return 0;
            }
            break;
        case WM_NCDESTROY:
            _host.Detach();
            if (_window.get() == hwnd)
                _window.release();
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            if (_returnFocus != nullptr && IsWindow(_returnFocus) != FALSE)
            {
                SetFocus(_returnFocus);
            }
            else if (_owner != nullptr && IsWindow(_owner) != FALSE)
            {
                SetFocus(_owner);
            }
            return DefWindowProcW(hwnd, message, wParam, lParam);
    }
    return DefWindowProcW(hwnd, message, wParam, lParam);
}

void CommandPaletteWindow::BuildUi()
{
    _rootStorage = std::make_unique<Panel>();
    _root        = _rootStorage.get();
    _searchField = _root->AddChild<TextField>();
    _searchField->SetPlaceholder(LoadStringResource(nullptr, IDS_COMMAND_PALETTE_SEARCH_CUE));
    _searchField->SetAccessibleName(LoadStringResource(nullptr, IDS_COMMAND_PALETTE_SEARCH_CUE));
    _searchField->SetOnTextChanged([this](std::wstring_view text)
    {
        _search = text;
        RebuildRows();
    });
    _grid = _root->AddChild<Grid>();
    _grid->SetDelegate(this);
    _grid->SetSelectionMode(GridSelectionMode::Single);
    _grid->SetHeaderHeightDip(0.0f);
    _grid->SetRowHeightDip(54.0f);
    _grid->SetLineClamp(2u);
    _emptyLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_DXUI_NO_MATCHES));
    _emptyLabel->SetAccessibleName(LoadStringResource(nullptr, IDS_DXUI_NO_MATCHES));
    _emptyLabel->SetVisible(false);
    _modelStorage = std::make_unique<PaletteGridModel>();
    _model        = _modelStorage.get();
    _grid->SetModel(_model);
    _host.SetRoot(std::move(_rootStorage));
}

void CommandPaletteWindow::RebuildRows()
{
    if (! _model || ! _grid)
        return;
    const auto filterStartedAt = std::chrono::steady_clock::now();
    std::vector<PaletteRow> rows;
    for (const auto& command : _catalog)
    {
        if (! command.paletteVisible)
        {
            continue;
        }
        std::wstring shortcuts;
        for (const auto& shortcut : command.shortcutTexts)
        {
            if (! shortcuts.empty())
                shortcuts.append(L"  ");
            shortcuts.append(L"‹");
            shortcuts.append(shortcut);
            shortcuts.append(L"›");
        }
        const bool titleMatch    = ContainsNoCase(command.displayName, _search);
        const bool shortcutMatch = ContainsNoCase(shortcuts, _search);
        if (! _search.empty() && ! titleMatch && ! shortcutMatch && ! ContainsNoCase(command.searchText, _search))
            continue;
        PaletteRow row;
        row.command = command;
        if (! ResolveCommandRuntimeStateFromWindow(_owner, _returnFocus, command.commandId, command.stateSource, row.runtimeState))
        {
            row.runtimeState.enabled = false;
        }
        row.commandText = command.displayName;
        if (! command.description.empty())
        {
            row.commandText.append(L"\n");
            row.commandText.append(command.description);
        }
        if (! row.runtimeState.enabled)
        {
            row.commandText.append(L"\n");
            row.commandText.append(LoadStringResource(nullptr, IDS_COMMAND_PALETTE_DISABLED));
        }
        row.shortcutText = std::move(shortcuts);
        row.iconText     = ResolveCommandVisualText(command.visualId, _host.HasFluentIconFont());
        row.rank         = _search.empty() ? 0u : StartsWithNoCase(command.displayName, _search) ? 0u : titleMatch ? 1u : shortcutMatch ? 2u : 3u;
        rows.push_back(std::move(row));
    }
    std::ranges::stable_sort(rows, [](const PaletteRow& left, const PaletteRow& right) noexcept { return left.rank < right.rank; });
    _model->SetRows(std::move(rows));
    _grid->NotifyDataChanged();
    _grid->SetVisible(_model->GetRowCount() != 0u);
    if (_emptyLabel != nullptr)
        _emptyLabel->SetVisible(_model->GetRowCount() == 0u);
    if (_model->GetRowCount() != 0u)
        static_cast<void>(_grid->RequestSelectRow(0u, 0u));
    _host.Invalidate();
    Debug::Perf::EmitDurationUs(L"app.command_palette.filter_batch_us", Debug::Perf::ElapsedUs(filterStartedAt), _catalog.size(), _model->GetRowCount());
}

void CommandPaletteWindow::Layout() noexcept
{
    if (! _root)
        return;
    const D2D1_RECT_F bounds = _host.GetClientBoundsDip();
    _root->SetBounds(bounds);
    constexpr float margin       = 14.0f;
    constexpr float searchHeight = 42.0f;
    _searchField->SetBounds(D2D1::RectF(margin, margin, bounds.right - margin, margin + searchHeight));
    _grid->SetBounds(D2D1::RectF(margin, margin + searchHeight + 10.0f, bounds.right - margin, bounds.bottom - margin));
    if (_emptyLabel != nullptr)
    {
        _emptyLabel->SetBounds(D2D1::RectF(margin + 18.0f, margin + searchHeight + 34.0f, bounds.right - margin - 18.0f, margin + searchHeight + 74.0f));
    }
}

void CommandPaletteWindow::ApplyTheme() noexcept
{
    _host.SetTheme(MakeAppThemeDxPalette(_theme));
    if (_window)
        ApplyWindowBackdropTheme(_window.get(), _theme, WindowBackdropTarget::Tool);
}

void CommandPaletteWindow::UpdateTheme(const AppTheme& theme) noexcept
{
    _theme = theme;
    ApplyTheme();
    _host.Invalidate();
}

void CommandPaletteWindow::Activate(size_t index) noexcept
{
    const PaletteRow* row = _model ? _model->GetRow(index) : nullptr;
    if (! row)
        return;
    CommandRuntimeState currentState{};
    if (! row->runtimeState.enabled ||
        ! ResolveCommandRuntimeStateFromWindow(_owner, _returnFocus, row->command.commandId, row->command.stateSource, currentState) ||
        ! currentState.enabled || ! CommandRuntimeIdentityMatches(row->runtimeState, currentState))
    {
        RebuildRows();
        return;
    }
    const std::wstring commandId = row->command.commandId;
    const HWND owner             = _owner;
    _window.reset();
    if (owner && IsWindow(owner) != FALSE)
        static_cast<void>(DispatchShortcutCommandFromWindow(owner, commandId));
}

#ifdef ENABLE_TESTS
bool CommandPaletteWindow::DebugSnapshot(CommandPaletteDebugSnapshot& out) const noexcept
{
    out = {};
    if (! _window || ! _model || ! _grid)
        return false;
    out.rowCount        = _model->GetRowCount();
    out.terminalContext = _terminalContext;
    out.searchText      = _search;
    const auto selected = _grid->GetPrimarySelectedRow();
    if (selected.has_value())
    {
        out.selectedRow = selected.value();
        if (const PaletteRow* row = _model->GetRow(selected.value()))
        {
            out.selectedCommandId     = row->command.commandId;
            out.selectedShortcutCount = row->command.shortcutTexts.size();
            out.selectedEnabled       = row->runtimeState.enabled;
        }
    }
    return true;
}
bool CommandPaletteWindow::DebugSetSearch(std::wstring_view text) noexcept
{
    if (! _searchField)
        return false;
    _searchField->SetText(std::wstring(text));
    // Programmatic test input does not rely on TextField's user-edit callback.
    // Keep the production filter state synchronized explicitly.
    if (_search != text)
    {
        _search = text;
        RebuildRows();
    }
    return true;
}
#endif
} // namespace

void ShowCommandPaletteWindow(HWND owner, const Common::Settings::ShortcutsSettings& shortcuts, bool terminalContext, const AppTheme& theme) noexcept
{
    if (g_palette && g_palette->GetHwnd() && IsWindow(g_palette->GetHwnd()) != FALSE)
    {
        SetForegroundWindow(g_palette->GetHwnd());
        return;
    }
    g_palette = std::make_unique<CommandPaletteWindow>();
    if (! g_palette->Create(owner, shortcuts, terminalContext, theme))
        g_palette.reset();
}

void UpdateCommandPaletteWindowTheme(const AppTheme& theme) noexcept
{
    if (g_palette && g_palette->GetHwnd())
        g_palette->UpdateTheme(theme);
}

HWND GetCommandPaletteWindowHandle() noexcept
{
    return g_palette ? g_palette->GetHwnd() : nullptr;
}

#ifdef ENABLE_TESTS
bool DebugGetCommandPaletteSnapshot(CommandPaletteDebugSnapshot& out) noexcept
{
    return g_palette && g_palette->DebugSnapshot(out);
}
bool DebugSetCommandPaletteSearch(std::wstring_view text) noexcept
{
    return g_palette && g_palette->DebugSetSearch(text);
}
#endif
