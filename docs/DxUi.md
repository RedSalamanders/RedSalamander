# DxUi Technical Guide

DxUi is RedSalamander's shared retained DirectX UI layer. It lives in
`Common/DxUi/` and is consumed by the main shell, Preferences, migrated dialogs,
viewer chrome, context menus, and component tests.

## Background

RedSalamander started with a mix of native Win32 controls, owner-draw controls,
and custom Direct2D surfaces. DxUi exists to move interactive UI into one shared
path with:

- Direct3D 11 / DXGI swap-chain presentation.
- Direct2D drawing and DirectWrite text.
- A retained control tree owned by `DxUi::WindowHost`.
- Shared theme tokens derived from app themes or viewer themes.
- Built-in keyboard focus, tab traversal, mnemonics, text input, tooltips,
  accessibility, pointer capture, animation, and high-DPI handling.

Use DxUi when a surface needs themed custom chrome, a migrated dialog body,
virtualized grid/tree behavior, or shared control behavior. Keep plain native
Win32 controls when the standard control is enough and no theme or behavior
parity issue is being solved.

## Main source files

| File | Responsibility |
| --- | --- |
| `Common/DxUi/DxUi.h` | Public control, theme, host, menu, grid, and tree API. |
| `Common/DxUi/DxUi.WindowHost.cpp` | HWND attachment, DirectX device/swap-chain lifecycle, message routing, focus, DPI, rendering, accessibility root. |
| `Common/DxUi/DxUi.Controls.cpp` | Common retained controls such as panels, labels, buttons, toggles, toolbars, tabs, and status strips. |
| `Common/DxUi/DxUi.TextInput.cpp` | `TextField` editing, caret, selection, clipboard, and text-input state sync. |
| `Common/DxUi/DxUi.NativeTextInput.cpp` and `DxUi.TextStoreACP.cpp` | Native text input, IME, TSF, caret windows, and UI Automation text events. |
| `Common/DxUi/DxUi.ComboBox.cpp` | Editable and non-editable combo boxes with popup, typeahead, scrolling, and text input. |
| `Common/DxUi/DxUi.Grid.cpp` | Virtualized grid, columns, sorting, selection, copy, checkboxes, badges, swatches, progress cells, scrollbars. |
| `Common/DxUi/DxUi.Tree.cpp` | Virtualized tree rows, expand/collapse, selection, typeahead, badges, scrollbars. |
| `Common/DxUi/DxUi.Menu.cpp` | DirectX context-menu flyouts and nested menu loop behavior. |
| `Common/DxUi/DxUi.Theme.cpp` | Default palette, viewer palette conversion, and resolved control visual styles. |
| `Common/DxUi/DxUiNativeMenuInterop.h` | Native menu bar interop backed by DxUi rendering and routing. |
| `Tests/DxUiTests/` | Focused deterministic tests and gallery/screenshot generation. |

## Host lifecycle

A DxUi surface has three pieces:

1. An HWND owned by the caller.
2. A `RedSalamander::DxUi::WindowHost` member whose lifetime covers the HWND
   attachment.
3. A retained root `Control` tree passed to `WindowHost::SetRoot(...)`.

Recommended lifecycle:

```cpp
using namespace RedSalamander::DxUi;

class ExamplePane final
{
public:
    bool Attach(HWND hwnd, const AppTheme& theme) noexcept
    {
        if (! hwnd || ! _host.Attach(hwnd))
        {
            return false;
        }

        auto root = std::make_unique<Panel>();
        root->SetBounds(D2D1::RectF(0.0f, 0.0f, 420.0f, 180.0f));

        auto* label = root->AddChild<Label>(L"Name");
        label->SetBounds(D2D1::RectF(16.0f, 16.0f, 140.0f, 40.0f));

        auto* field = root->AddChild<TextField>();
        field->SetPlaceholder(L"Connection name");
        field->SetBounds(D2D1::RectF(16.0f, 44.0f, 280.0f, 76.0f));
        field->SetOnSubmitted([this]() { Save(); });
        _nameField = field;

        auto* ok = root->AddChild<Button>(L"OK");
        ok->SetBounds(D2D1::RectF(300.0f, 132.0f, 380.0f, 164.0f));
        ok->SetOnClick([this]() { Save(); });
        _okButton = ok;

        _host.SetTheme(MakeAppThemeDxPalette(theme));
        _host.SetRoot(std::move(root));
        _host.SetDefaultButton(_okButton);
        return true;
    }

    LRESULT WindowProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept
    {
        return _host.HandleMessage(hwnd, msg, wp, lp, handled);
    }

    void ApplyTheme(const AppTheme& theme) noexcept
    {
        _host.SetTheme(MakeAppThemeDxPalette(theme));
    }

    void Detach() noexcept
    {
        _host.Detach();
        _nameField = nullptr;
        _okButton = nullptr;
    }

private:
    void Save();

    WindowHost _host;
    TextField* _nameField = nullptr;
    Button* _okButton = nullptr;
};
```

Key rules:

- Store `WindowHost` in stable owner storage. Do not store subclass refdata that
  can be invalidated by moving the owner.
- Forward messages to `WindowHost::HandleMessage(...)` before local fallback
  handling for input, focus, paint, size, DPI, accessibility, and text input.
- Detach explicitly before destroying owner state or bulk-destroying child HWNDs.
- Use `SetDefaultButton(...)`, `SetCancelButton(...)`, `SetOnEscape(...)`, and
  `SetOnTabBoundary(...)` instead of reimplementing dialog keyboard behavior.
- Call `SetTheme(...)` whenever the app or viewer theme changes.

## Backgrounds and theme surfaces

DxUi paints its own surface. The host HWND should not rely on a Win32 class
background brush for the visible DxUi area. For app surfaces, derive the palette
through `MakeAppThemeDxPalette(...)` in `RedSalamander/DxUiThemePalette.h`.
For viewer plugins, use `MakeThemePaletteFromViewerTheme(...)`.

Common palette fields:

- `windowBackground`: the outer window background.
- `surfaceBackground`: cards, dialog bodies, and hosted content surfaces.
- `overlayBackground`: menus, popups, and tool windows.
- `headerBackground`, `headerHovered`, `headerPressed`: menu/header chrome.
- `text`, `subduedText`, `disabledText`: text hierarchy.
- `selectionFill`, `selectionText`, `focusStroke`: selection and keyboard focus.
- `inputFill`, `inputBorder`: text fields and combo boxes.
- `scrollbarTrack`, `scrollbarThumb`, `scrollbarThumbHot`: shared scrollbars.

Use the optional `surfaceBackground` argument when a host sits on a specific
parent surface:

```cpp
const auto palette = MakeAppThemeDxPalette(theme, theme.windowBackground);
_host.SetTheme(palette);
```

Whole-window or popup backdrops may request a Windows 11 DWM material:

```cpp
if (! theme.highContrast)
{
    _host.SetSystemBackdrop(WindowHost::BackdropType::Mica);
}
```

Keep high-contrast and reduced-motion behavior in the palette. Avoid ad-hoc
colors in paint code unless a new named visual token is being introduced in
`DxUi.Theme.cpp` with focused tests.

## Message routing pattern

For a direct owner WndProc:

```cpp
case WM_NCDESTROY:
    state->host.Detach();
    break;

default:
{
    bool dxHandled = false;
    const LRESULT dxResult = state->host.HandleMessage(hwnd, msg, wp, lp, dxHandled);
    if (dxHandled)
    {
        return dxResult;
    }
    break;
}
```

For a child host installed into an existing Win32 shell:

- Install the subclass/hook only after the owner storage is stable.
- Route `WM_NCDESTROY` through detach and hook cleanup even if DxUi does not
  consume the message.
- If the wrapper window class is `Static`, include `SS_NOTIFY` so pointer input
  reaches the host.
- Child `WindowHost` surfaces should return `MA_ACTIVATE` for `WM_MOUSEACTIVATE`
  through the shared host path so the first click activates the control.

## Controls

Common controls are declared in `Common/DxUi/DxUi.h`:

- Layout and containers: `Panel`, `StackPanel`, `ScrollPanel`, `PageHost`,
  `CardPanel`, `TabControl`, `Toolbar`, `StatusStrip`.
- Text and commands: `Label`, `Button`, `Toggle`, `Checkbox`, `RadioButtons`,
  `ProgressBar`, `Slider`, `ColorSwatch`.
- Inputs: `TextField`, `ComboBox`.
- Data surfaces: `Grid`, `Tree`.
- Menus: `MenuBar`, `ContextMenu`, `NativeMenuBarHost`.

Most controls use DIPs, not pixels. Convert only at the host boundary with
`PixelsToDip(...)`, `DipsToPixels(...)`, or `ScreenPointToDipPoint(...)`.

## Example: editable combo in a viewer

Viewer plugins often host a compact DxUi combo inside a legacy viewer shell:

```cpp
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::MakeDefaultThemePalette;
using RedSalamander::DxUi::MakeThemePaletteFromViewerTheme;

bool ViewerPane::AttachFileCombo(HWND hwnd) noexcept
{
    if (! _fileComboHost.Attach(hwnd))
    {
        return false;
    }

    auto combo = std::make_unique<ComboBox>();
    combo->SetVariant(ComboBoxVariant::Edit);
    combo->SetEditable(true);
    combo->SetItems(BuildFileItems());
    combo->SetOnSelectionChanged([this](size_t index) { SelectFile(index); });
    combo->SetOnSubmitted([this]() { OpenTypedFile(); });

    _fileComboControl = combo.get();
    _fileComboHost.SetRoot(std::move(combo));
    _fileComboHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme)
                                      : MakeDefaultThemePalette(false));
    return true;
}
```

Use `Common/ViewerFileComboHost.h` for shared escape/tab focus return behavior
when embedding a viewer file combo in an existing viewer window.

## Example: virtualized grid

`DxUi::Grid` reads from an `IDxGridModel` and reports actions to an
`IDxGridDelegate`. The model is non-owning and must outlive the grid.

```cpp
class RowModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    size_t GetColumnCount() const noexcept override { return _columns.size(); }
    size_t GetRowCount() const noexcept override { return _rows.size(); }

    RedSalamander::DxUi::GridColumnDesc GetColumn(size_t column) const override
    {
        return _columns.at(column);
    }

    void GetCellData(size_t row, size_t column, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        outCell.text = _rows.at(row).at(column);
    }

    std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId >= _rows.size())
        {
            return std::nullopt;
        }
        return static_cast<size_t>(rowId);
    }

private:
    std::vector<RedSalamander::DxUi::GridColumnDesc> _columns;
    std::vector<std::array<std::wstring, 3>> _rows;
};

auto* grid = root->AddChild<RedSalamander::DxUi::Grid>();
grid->SetModel(&_model);
grid->SetDelegate(this);
grid->SetSelectionMode(RedSalamander::DxUi::GridSelectionMode::Extended);
grid->SetEmptyStateText(L"No items");
_grid = grid;
```

When the data changes, update the model first, then call:

```cpp
_grid->NotifyDataChanged();
_host.Invalidate();
```

Do not mutate the model from a worker thread while the grid is reading it on the
UI thread. Post a payload to the UI owner, then update the model and notify the
grid from the UI thread.

## Text input and IME

`TextField` and editable `ComboBox` use the native retained DxUi text-input path
owned by `WindowHost`. Production code should use:

- `WindowHost::CommitFocusedTextInput()` before reading focused edit contents in
  command handlers.
- `WindowHost::TryReadTextInputState(...)` for diagnostics and tests.
- `WindowHost::SyncTextInput(...)` after code mutates retained text while the
  control is focused.

Do not create hidden edit controls, hidden RichEdit controls, or bridge-specific
messages for new production code.

## Accessibility

`WindowHost` answers `WM_GETOBJECT` and creates a retained UI Automation root.
Interactive controls should expose names, roles, enabled state, selection state,
toggle state, and value/text patterns through the shared providers. When adding
or changing a control, add focused coverage in `Tests/DxUiTests`, usually in
`DxUiTests.Accessibility.cpp`.

## Performance and rendering

DxUi is used on hot UI paths. Keep these constraints in mind:

- Prefer one shared host for a migrated dialog/page over one swap-chain host per
  control.
- Hidden or minimized hosts should not keep resizing swap chains, invalidating,
  or presenting.
- Device resources are shared on the UI thread and must not be touched from
  worker threads.
- Grid and tree implementations should do visible-work only. Use
  `GridVisibleWorkMetrics` and focused tests when changing virtualization.
- Avoid per-cell/per-frame allocations in paint paths. Cache DirectWrite layouts
  only when the cache invalidation rules are clear.

## Tests and galleries

Build focused tests:

```powershell
.\build.ps1 -ProjectName DxUiTests
```

Run focused suites:

```powershell
.\build\x64\Debug\DxUiTests.exe --suite=WindowHost
.\build\x64\Debug\DxUiTests.exe --suite=Control
.\build\x64\Debug\DxUiTests.exe --suite=ComboBox
.\build\x64\Debug\DxUiTests.exe --suite=Grid
.\build\x64\Debug\DxUiTests.exe --suite=Tree
.\build\x64\Debug\DxUiTests.exe --suite=NativeTextInput
.\build\x64\Debug\DxUiTests.exe --suite=Accessibility
```

Refresh gallery images used by [Themes.md](Themes.md):

```powershell
.\build\x64\Debug\DxUiTests.exe --suite=Gallery --gallery-output-directory=docs\res
```

## Common pitfalls

- Forgetting to call `Detach()` before owner state is destroyed.
- Keeping a raw `WindowHost*` in window properties after the owner can move or
  be destroyed.
- Reintroducing hidden native edit/combo controls behind retained DxUi controls.
- Painting direct color choices instead of adding/resolving named theme tokens.
- Reading or mutating grid/tree models from a worker thread.
- Swallowing Tab/Escape/Enter locally instead of using host callbacks and
  control-level submission/cancel behavior.
- Forgetting accessibility and keyboard tests for a visual-only migration.

## More detail

For normative behavior, read:

- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/UI/UI_VisualStyle.md`
- `Specs/UI/UI_NavigationView.md`
- `Specs/Testing/Testing_PerformanceValidation.md`
