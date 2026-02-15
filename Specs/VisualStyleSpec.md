# Visual Style Specification

This document defines shared visual tokens and patterns used across RedSalamander UI components.

## Rounded Corners (Selection/Focus)

- **Standard corner radius**: `2 DIP` (both X/Y).
- **Applies to**:
  - `FolderView` item **selected** and **hover** backgrounds.
  - `FolderView` item **focus** border.
  - `ColorTextView` (Monitor) text **selection** highlight rectangles.
  - `NavigationView` keyboard **focus** ring rectangles.
  - Future **marquee selection** rectangles (if implemented).
- **Implementation guidance**:
  - Direct2D: use `FillRoundedRectangle` / `DrawRoundedRectangle` with `radiusX = radiusY = 2.0f`.
  - Clamp per-rectangle radius to avoid artifacts on very small rects (do not exceed `min(width, height) / 2`).

## NavigationView Breadcrumb Hover/Pressed

- Breadcrumb segment/separator hover (and separator pressed) backgrounds use an **inset rounded rectangle**:
  - Inset: `1 DIP` (horizontal and vertical)
  - Corner radius: `2 DIP`
- Breadcrumb segment text has **symmetric left/right padding** so hover backgrounds show equal gaps on both sides.

## NavigationView Edit Mode (Address Bar)

- **Edit underline**: `2 DIP` height line at the bottom of the editable path field (does not extend under the close button).
- **Close button** (right side):
  - Width: `24 DIP`
  - Icon: centered `X`, `10 DIP` size (5 DIP half-size), `1.5 DIP` stroke
  - Hover background: same **inset rounded rectangle** style as breadcrumb hover (1 DIP inset, 2 DIP corner radius).
- **Edit text padding**: `6 DIP` left and right inside the edit field (to match breadcrumb padding).
- **Autosuggest popup**:
  - Opens under the edit field while typing.
  - Shows up to `10` items.
    - If there are more than `10` matches, the last item is a disabled `...` indicator.
  - Highlights the typed substring inside each suggestion item (accent + semibold).
  - Uses the same `2 DIP` rounded-corner selection/hover highlight style as other list selections.

## NavigationView Active/Inactive Pane

NavigationView reflects the **focused pane** (not only whether NavigationView itself has keyboard focus):

- **Active pane**: NavigationView uses the normal background and an accent bottom border.
- **Inactive pane**: NavigationView uses a noticeably dimmed palette (including icons) while keeping text readable.
- **Theme-aware**: Works consistently in Light/Dark/Rainbow by deriving from `AppTheme.navigationView` colors.

## FolderWindow Splitter Grip

- **Splitter width**: `6 DIP` (existing).
- **Grip**: a subtle centered "handle" inside the splitter to indicate it is draggable.
  - Shape: `3` vertically stacked dots (small squares are acceptable).
  - Dot size: `2 DIP` square.
  - Gap between dots: `2 DIP`.
  - Alignment: centered within the splitter rectangle (both axes).
  - Color: slightly higher contrast than the splitter fill (derive from `AppTheme.menu.separator` blended toward `AppTheme.menu.text`; prefer `AppTheme.menu.text` in High Contrast mode).
- **Hit target**: purely visual; the full splitter rect remains the draggable region.

## Implementation Touchpoints (Code)

- `RedSalamander/FolderWindow.cpp`: `FolderWindow::OnPaint()` (splitter fill + grip).
- `RedSalamander/FolderView.cpp`: `FolderView::DrawItem()` (selection/hover fill + focus border).
- `RedSalamanderMonitor/ColorTextView.cpp`: `ColorTextView::DrawSelection()` (selection highlight).
- `RedSalamander/NavigationView.cpp`: focus ring drawing in `Render*Section()` paths.

## Menus (Popup Submenus)

- Popup menu items that open a submenu render a **thin chevron-right** indicator (Segoe Fluent Icons preferred; fallback to a Unicode chevron).
- **Color**: submenu chevrons are rendered using the theme’s *secondary* color (e.g. `MenuTheme.shortcutText` / selected variant), not the main label text color.
- **Color (glyph icons)**: menu item glyph icons (Segoe Fluent Icons) are rendered using the same color as the menu label text (`MenuTheme.text` / `MenuTheme.selectionText` / `MenuTheme.disabledText`) so they remain legible in dark theme.
- **Size (glyph icons)**: Segoe Fluent icon fonts default to `15 DIP` (see `FluentIcons::kDefaultSizeDip`) and scale with DPI.
- **Owner-draw arrow gotcha (MUST)**: when using `MFT_OWNERDRAW`, Windows may still paint the default submenu arrow *after* `WM_DRAWITEM`. To ensure only the custom chevron is visible, the draw handler must clip out the arrow area after drawing it:
  - Set an initial clip region to the item rect (so exclusions don’t leak).
  - Compute an arrow rect at the right edge using `max(customArrowAreaWidth, GetSystemMetricsForDpi(SM_CXMENUCHECK, dpi))`.
  - Call `ExcludeClipRect(hdc, arrowRect.left, arrowRect.top, arrowRect.right, arrowRect.bottom)` before returning from `WM_DRAWITEM`.

## Themed Win32 Controls (Dark Mode + Scrollbars)

Windows common controls do not automatically inherit dark-mode styling in all cases. Any new UI that uses Win32 common controls (especially controls with scrollbars) MUST apply theme plumbing explicitly.

- **Apply `SetWindowTheme` per control** based on the active theme:
  - High contrast: `SetWindowTheme(hwnd, L"", nullptr)` (let the system fully own colors)
  - Dark mode: `SetWindowTheme(hwnd, L"DarkMode_Explorer", nullptr)`
  - Light/system: `SetWindowTheme(hwnd, L"Explorer", nullptr)`
- **Forward `WM_THEMECHANGED`**: after changing the theme or calling `SetWindowTheme`, also send `SendMessageW(hwnd, WM_THEMECHANGED, 0, 0)` so the control refreshes its non-client/scrollbar theming.
- **Controls that require explicit theming** (minimum set):
  - Scrollbar-owning controls: `EDIT` / `RICHEDIT` (`MSFTEDIT_CLASS`), `LISTVIEW`, `TREEVIEW`, `LISTBOX`
  - Composite controls: `COMBOBOX` **and** its child windows (`COMBOBOXINFO.hwndList`, `COMBOBOXINFO.hwndItem`)
  - `LISTVIEW` header (`ListView_GetHeader(hwndListView)`)
- **Dark combobox/listbox backgrounds**: for dropdowns that still paint with a light background, handle `WM_CTLCOLORLISTBOX` / `WM_CTLCOLOREDIT` / `WM_CTLCOLORSTATIC` and return a theme brush + `SetTextColor/SetBkColor` (skip this in high contrast).

## Themed Title Bars (Top-Level Windows)

Any new top-level window that is theme-aware (dialogs included unless explicitly mentioned) MUST apply title-bar styling:

- **Dark mode**: use `DwmSetWindowAttribute` with immersive dark mode attributes (`DWMWA_USE_IMMERSIVE_DARK_MODE` 19/20) to match app theme.
- **Rainbow mode**: set caption/border/text colors (`DWMWA_BORDER_COLOR` 34, `DWMWA_CAPTION_COLOR` 35, `DWMWA_TEXT_COLOR` 36) to an accent-derived color and a contrasting text color.
- Apply these after the window handle exists and re-apply on theme changes.

## FolderView Focused Pane Indicator

RedSalamander is a dual-pane UI. Users must be able to identify the **focused pane** from the file list alone.

- **Focused pane**: the `Current item` draws **border + background**.
- **Unfocused pane**: the `Current item` draws **border only** (no background fill) and uses a **thinner + dimmer** focus stroke (reduced opacity).
- `Selected items` remain visually distinct in both panes; the unfocused-pane selection palette comes from theme tokens (`FolderViewTheme.itemBackgroundSelectedInactive` / `FolderViewTheme.textSelectedInactive`) and must not be hard-coded.

See `Specs/CommandMenuKeyboardSpec.md` for the behavioral definition of focused/active pane.

## Plugin Configuration / Plugins Manager Dialogs

### Plugin Configuration Dialog

- **Per-field layout**: left label column + right control column.
- **Description text**: render the schema `description` **under the field name** (left-aligned) using an **italic** font.
- **Default/min/max text**: render defaults and numeric ranges (`Default` / `Min` / `Max`) in **italic** under the description.
- **Two-value options**: options with exactly 2 choices render as a **custom toggle** (labels left/right, selected label bold, accent knob indicates active side).
- **Multi-value options**: options with more than 2 choices render as a **themed combobox** (not a vertical radio list).
- **Numeric inputs**: use a **narrower** edit width than free-form text inputs.
- **Input styling**: use a **rounded-rectangle** frame; avoid `WS_EX_CLIENTEDGE` 3D borders when not in high contrast.
- **Sizing**: the dialog is **vertically resizable**; the configuration panel scrolls when content exceeds the available height.
- **Keyboard**: tab navigation works through all configuration controls (including when the panel is scrollable).
- **Focus indication**: custom-drawn buttons/toggles use a **subtle** keyboard focus outline (avoid high-contrast XOR focus rectangles).
- **Borders**: avoid a global “outer border” around the entire configuration panel.

### Plugins Manager Dialog

- **Buttons**: dialog push buttons should be **theme-aware** (custom-drawn) so they match Light/Dark palettes (skip custom drawing in high contrast).
