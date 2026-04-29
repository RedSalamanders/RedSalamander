# DxUI Migration Audit — Comprehensive Codebase Analysis

**Author:** Ripley (Lead / Reviewer)  
**Date:** 2025-07-25  
**Status:** Analysis Complete  
**Impact:** Architecture / Migration Planning  

---

## 1. Current DxUI Control Inventory (18 Controls + Infrastructure)

The DxUI framework lives in `Common/DxUi/` and is architecturally mature — 13 implementation files, full UI Automation accessibility, device-loss recovery, per-monitor DPI, theme system.

### Infrastructure
| Component | Description |
|-----------|-------------|
| **WindowHost** | Core D2D/D3D rendering host, focus/animation/tooltip management, text input bridge, accessibility root |
| **Control** | Abstract base — paint, hit test, events, focus, mnemonic |
| **ThemePalette** | Full color palette with dark/high-contrast/reduced-motion/accent support |
| **GridSelectionModel** | Selection state manager (single, range, multi) |

### Containers & Layout
| Control | Description |
|---------|-------------|
| **Panel** | Base container with child management |
| **CardPanel** | Rounded-corner grouped container |
| **StackPanel** | Auto layout (vertical/horizontal), gap, padding |
| **ScrollPanel** | Scrollable container with vertical scrollbar |
| **PopupLayer** | Overlay rendering for dropdowns/popups |

### Display Controls
| Control | Description |
|---------|-------------|
| **Label** | Text display, font roles, mnemonic, multiline, alignment |
| **StatusStrip** | Status bar text (inherits Label) |
| **TooltipLayer** | Tooltip rendering at arbitrary position |
| **ColorSwatch** | Color display/selection |

### Input Controls
| Control | Description |
|---------|-------------|
| **Button** | Click, primary/secondary style, mnemonic, keyboard invoke |
| **Toggle** | Two-state switch with state labels |
| **Checkbox** | Standard checkbox (inherits Toggle) |
| **TextField** | Single/multi-line text input, selection, clipboard, IME, masked input |
| **ComboBox** | Dropdown with 3 variants (Window/Modern/Edit), typeahead, editable |

### Data Controls
| Control | Description |
|---------|-------------|
| **Tree** | Hierarchical tree with model/delegate, icons, badges, typeahead |
| **Grid** | High-performance table with sorting, grouping, column resize/reorder, multi-select, cell types |

### Model/Delegate Interfaces
- **IDxGridModel** / **IDxGridDelegate** — Grid data source and interaction callbacks
- **IDxTreeModel** / **IDxTreeDelegate** — Tree data source and interaction callbacks

---

## 2. Common Controls (comctrl32) Dependencies

### Summary: 257 control usages across 33 source files

| Component | Control Type | File(s) | Complexity | DxUI Equivalent Needed |
|-----------|-------------|---------|------------|----------------------|
| **ConnectionManagerDialog** | ListView | ConnectionManagerDialog.cpp (35+ macro ops, lines 3324-3539) | **HARD** | Grid (already exists) — needs card-based item rendering mode |
| **Preferences.Plugins** | ListView + Header | Preferences.Plugins.cpp (13 LV + 5 TV + 2 HD macros) | **MEDIUM** | Grid (exists) — extension mapping list |
| **Preferences.Dialog** | TreeView | Preferences.Dialog.cpp (6 TV macros, lines 3286-4398) | **MEDIUM** | Tree (exists) — category navigation |
| **Preferences.Themes** | TreeView | Preferences.Themes.cpp (4 TV color ops) | **MEDIUM** | Tree (exists) — theme color list |
| **ThemedControls** | ListView + Header | ThemedControls.cpp (6 LV + 3 HD, lines 2643-2816) | **MEDIUM** | Grid (exists) — centralized theming point, becomes unnecessary with DxUI |
| **RedSalamanderMonitor** | Toolbar | RedSalamanderMonitor.cpp (9 TB_ messages, line 1022+) | **MEDIUM** | **NEW: DxUI Toolbar** needed |
| **RedSalamanderMonitor** | StatusBar | RedSalamanderMonitor.cpp (7 SB_ messages, line 1314+) | **MEDIUM** | StatusStrip exists but needs multi-section support → **NEW: DxUI StatusBar** |
| **ViewerSpace** | Tooltip | ViewerSpace.cpp (TTM_ messages, lines 4698-4805) | **MEDIUM** | TooltipLayer exists — needs tracking tooltip mode |
| **CompareDirectoriesWindow** | ProgressBar | CompareDirectoriesWindow.cpp (ICC_PROGRESS_CLASS, line 1546) | **SIMPLE** | **NEW: DxUI ProgressBar** needed |
| **ViewerText** | ScrollBar (SB_) | ViewerText.Text.cpp (31 cases), ViewerText.Hex.cpp (16 cases) | **SIMPLE** | Scrollbar exists in Grid/Tree/ScrollPanel — needs standalone extraction |
| **ColorTextView** | ScrollBar (SB_) | ColorTextView.cpp (24 SB_ cases, lines 865-991) | **SIMPLE** | Same standalone scrollbar needed |
| **ViewerImgRaw** | ScrollBar (SB_) | ViewerImgRaw.cpp (24 cases) | **SIMPLE** | Same standalone scrollbar |
| **ViewerPE** | ScrollBar (SB_) | ViewerPE.cpp (10 cases) | **SIMPLE** | Same standalone scrollbar |
| **ManagePluginsDialog** | ListView | ManagePluginsDialog.cpp (12 cases) | **MEDIUM** | Grid (exists) |
| **CompareDirectoriesWindow.Options** | ScrollBar | CompareDirectoriesWindow.Options.cpp (11 cases) | **SIMPLE** | Standalone scrollbar |
| **FolderWindow.FileOperations.Popup** | ScrollBar | FolderWindow.FileOperations.Popup.cpp (15 cases) | **SIMPLE** | Standalone scrollbar |

### Control Type Totals

| Win32 Control | Occurrences | Components Using It | DxUI Status |
|---------------|-------------|--------------------|----|
| ScrollBar (SB_) | 93 | 8 files across all viewer plugins + dialogs | **Partial** — embedded in Grid/Tree/ScrollPanel, needs standalone extraction |
| ListView (LVM_/ListView_) | 73 | ConnectionManager, Preferences, ThemedControls, ManagePlugins | **Covered** — Grid handles these use cases |
| Header (HDM_/Header_) | 34 | Tied to ListView usage | **Covered** — Grid has built-in header |
| TreeView (TVM_/TreeView_) | 15 | Preferences category nav, theme color list | **Covered** — Tree handles these |
| Toolbar (TB_) | 11 | RedSalamanderMonitor only | **MISSING** |
| Tooltip (TTM_) | 10 | ViewerSpace only | **Partial** — TooltipLayer exists, needs tracking mode |
| StatusBar (SB_SET*) | 8 | RedSalamanderMonitor only | **Partial** — StatusStrip exists but single-section only |
| ProgressBar (PBM_) | 1 | CompareDirectoriesWindow | **MISSING** |

---

## 3. Win32 GDI/User32 Dependencies

### Summary: ~300+ GDI API calls, ~40 WM_PAINT handlers, 75+ custom WndProc implementations

| Component | Win32 API | File(s) | Complexity | DxUI Approach |
|-----------|-----------|---------|------------|--------------|
| **Custom Menu System** | GetDC, GetTextExtentPoint32W, CreatePen, FillRect, MoveToEx, LineTo, CreateRectRgnIndirect | FolderView.Menus.cpp, NavigationView.Menus.cpp, RedSalamander.cpp, CompareDirectoriesWindow.Menu.cpp, ViewerText.MenuTheme.cpp, ViewerImgRaw.cpp, ViewerSpace.cpp, ViewerWeb.cpp | **HARD** | **NEW: DxUI MenuBar/ContextMenu** — owner-drawn menus are the #1 GDI hotspot (~15 locations, ~500 lines) |
| **FunctionBar** | BeginPaint, FillRect, GetDC, GetTextExtentPoint32W | FunctionBar.cpp (lines 367-519) | **MEDIUM** | **NEW: DxUI FunctionBar** component |
| **StatusBar Rendering** | GetDC, GetTextExtentPoint32W, BeginPaint | FolderWindow.StatusBar.cpp (lines 39-484) | **MEDIUM** | **NEW: DxUI StatusBar** with multi-section, measurement |
| **Themed Controls Painting** | BeginPaint, CreateSolidBrush, FillRect, CreatePen, RoundRect, GetTextMetricsW | ThemedControls.cpp (lines 1057-2816) | **MEDIUM** | DxUI replaces this entirely — ThemedControls.cpp becomes dead code |
| **ConnectionManager Dialog** | BeginPaint, CreateCompatibleDC, CreateCompatibleBitmap, BitBlt | ConnectionManagerDialog.cpp (lines 4923-4949) | **HARD** | DxUI Panel-based dialog with Grid |
| **Compare Directories** | BeginPaint, FillRect | CompareDirectoriesWindow.cpp (lines 1137-1149) | **MEDIUM** | DxUI Panel layout |
| **Preferences Dialog** | BeginPaint, CreateCompatibleDC, CreateCompatibleBitmap | Preferences.Dialog.cpp (lines 3412-3428), Preferences.Internal.cpp | **MEDIUM** | DxUI (partially migrated already) |
| **NavigationView** | BeginPaint | NavigationView.cpp, NavigationView.Edit.cpp, NavigationView.FullPathPopup.cpp | **HARD** | DxUI with embedded text editing |
| **Viewer Plugins (6)** | BeginPaint, CreatePen, GetDC, GetTextMetricsW | ViewerText, ViewerWeb, ViewerVLC, ViewerImgRaw, ViewerSpace, ViewerPE | **MEDIUM** | Each viewer needs a DxUI rendering surface or wrapper |
| **Icon Premultiplication** | GetDC, CreateCompatibleDC, CreateDIBSection, SelectObject, DrawIconEx | FolderView.Rendering.cpp (lines 799-852), IconCache.cpp | **SIMPLE** | Keep — GDI-to-D2D bridge is required for shell icon conversion |
| **Alert Overlay** | BeginPaint | AlertOverlayWindow.cpp | **SIMPLE** | DxUI overlay component |
| **Monitor Main Window** | BeginPaint | RedSalamanderMonitor.cpp (line 2297) | **SIMPLE** | DxUI shell |
| **ColorTextView Find Panel** | WndProc with child edit controls | ColorTextView.cpp (find panel: label, edit, checkboxes) | **MEDIUM** | DxUI Panel + TextField + Checkbox |

### GDI API Category Totals

| API Category | Count | Migration Priority |
|--|--|--|
| WM_PAINT / BeginPaint | ~40 handlers | P1 — each becomes a DxUI Paint() |
| Text Measurement (GetTextExtentPoint32W, GetTextMetricsW) | ~35 calls | P1 — use DxUI DirectWrite text layout |
| GDI Drawing (CreatePen, FillRect, MoveToEx, LineTo, RoundRect) | ~25 calls | P1 — use D2D geometry/brush APIs |
| Double Buffering (CreateCompatibleDC, CreateCompatibleBitmap, BitBlt) | ~15 calls | P0 — DxUI handles this automatically |
| Window Class Registration (RegisterClassExW, WNDCLASSEXW) | ~60+ registrations | Incremental — each becomes a DxUI component class |

### Custom WndProc Implementations (75+)

**Major Application Windows (9):** FolderWindow, FolderView, NavigationView, FunctionBar, CompareDirectoriesWindow, FindFilesWindow, ShortcutsWindow, SplashScreen, RedSalamander main

**Plugin Viewer Windows (13):** ViewerWeb, ViewerVLC (5 WndProcs!), ViewerText (3), ViewerPE, ViewerImgRaw, ViewerSqlite, ViewerSpace

**Monitor Application (3):** RedSalamanderMonitor, ColorTextView (3 WndProcs)

**Dialog/Host Processors (20+):** Preferences page hosts, ConnectionManager hosts, CompareOptions hosts, Keyboard/Plugin config hosts

**Themed Control Processors (7):** ThemedButtonHoverWndProc, ModernComboListWndProc, ModernComboPopupWndProc, ModernComboWndProc, ListViewHeaderWndProc, InputControlWndProc, InputFrameWndProc

### HWND Member Variables (90+)

**RedSalamander:** ~55 HWNDs across FolderWindow (5), FolderView (2), NavigationView (7), FunctionBar (1), CompareDirectoriesWindow (10), Preferences pages (16+), AlertOverlay (4), AnimationDispatcher (1), FileOperations (2)

**Plugins:** ~20 HWNDs across 7 viewer plugins (ViewerVLC has 5 alone)

**Monitor:** ~7 HWNDs (ColorTextView: 6, Monitor: 1)

**Common/DxUI:** 2 HWNDs (host window + text input bridge — these stay)

---

## 4. DxUI Framework Gaps

### P0: Blocking Multiple Components

| Missing Control/Feature | Blocks | Effort |
|------------------------|--------|--------|
| **Standalone Scrollbar** | All 6 viewer plugins, ColorTextView, CompareDirectoriesWindow, FileOperations popup (93 SB_ occurrences, 8 files) | **MEDIUM** — logic exists in Grid/Tree/ScrollPanel, needs extraction into reusable `DxScrollbar` |
| **Menu System (MenuBar + ContextMenu)** | FolderView menus, NavigationView menus, main app menu, CompareDirectoriesWindow menu, all 6 viewer plugin menus (~15 locations, ~500 lines of GDI menu drawing) | **HARD** — owner-drawn menus are the single largest GDI dependency; needs popup window hosting, keyboard nav, icons, separators, shortcuts, submenus |

### P1: Needed for Specific Component Migration

| Missing Control/Feature | Blocks | Effort |
|------------------------|--------|--------|
| **Toolbar** | RedSalamanderMonitor toolbar (11 TB_ messages, icon/button/sizing) | **MEDIUM** — horizontal button strip with icon support, separator, auto-size |
| **StatusBar (multi-section)** | RedSalamanderMonitor status bar (5 sections), FolderWindow status bar | **MEDIUM** — StatusStrip exists but only single text; needs N sections with auto-layout |
| **ProgressBar** | CompareDirectoriesWindow scan progress | **SIMPLE** — horizontal bar with determinate/indeterminate modes; Grid already has Marquee/Spinner cell types to draw from |
| **RadioButton** | Preferences pages with mutually-exclusive options | **SIMPLE** — Toggle variant with group exclusion logic |
| **TabControl** | Connection manager, viewer plugin multi-tab (if not using Tree for navigation) | **MEDIUM** — tab strip with content switching; StackPanel + button strip could approximate |
| **Tracking Tooltip** | ViewerSpace hover metadata (TTM_TRACKPOSITION, TOOLINFOW) | **SIMPLE** — TooltipLayer exists, needs `SetOrigin()` to follow mouse and tracking activate/deactivate |

### P2: Nice-to-Have / Workaround Available

| Missing Control/Feature | Notes | Effort |
|------------------------|-------|--------|
| **FunctionBar** | Specific to RedSalamander; could be custom DxUI control | **MEDIUM** — 12-key strip with hover, themed backgrounds |
| **SplitContainer / Splitter** | FolderWindow dual-pane, CompareDirectoriesWindow split | **MEDIUM** — draggable divider between two panels |
| **Dialog / Modal Host** | Preferences, ConnectionManager, ManagePlugins, Find files | **HARD** — needs title bar, close button, modal overlay, keyboard trap, ESC handling |
| **RichTextView** | ColorTextView D2D text viewer (~200KB) is already custom D2D; not a DxUI control per se | **ALREADY D2D** — doesn't need DxUI migration, just eventual integration |
| **NumericUpDown / Spinner** | Preferences numeric fields | **SIMPLE** — TextField + increment/decrement buttons |
| **Hyperlink Label** | About dialog, preference explanatory text | **SIMPLE** — Label variant with click handler and underline style |

---

## 5. Recommended Migration Roadmap

### Phase 1: Extract & Extend Core DxUI Primitives (Foundation)
**Priority: P0 | Effort: 2-3 weeks | Unblocks: Everything**

1. **Extract Standalone DxScrollbar** from Grid/Tree/ScrollPanel into reusable `DxScrollbar` control
   - Unblocks all 6 viewer plugins + ColorTextView + 3 dialog scrollbar usages
   - ~200 lines to extract, API: SetRange/SetPosition/OnScroll callback
   
2. **Extend StatusStrip → DxStatusBar** with multi-section support
   - Unblocks Monitor status bar + FolderWindow status bar
   - Add SetParts/SetText per section, auto-layout

3. **Add DxProgressBar** (determinate + indeterminate/marquee)
   - Unblocks CompareDirectoriesWindow
   - Borrow animation pattern from Grid's Marquee cell type

### Phase 2: Menu System (Highest GDI Impact)
**Priority: P0 | Effort: 3-4 weeks | Eliminates ~500 lines of GDI code across 15 files**

4. **Build DxUI ContextMenu** — popup menu with items, icons, separators, shortcuts, keyboard navigation
   - Start with FolderView context menu (most complex, most user-facing)
   - Reuse PopupLayer + vertical StackPanel + keyboard handling from ComboBox dropdown

5. **Build DxUI MenuBar** — horizontal menu strip with dropdown activation
   - For main application menu and CompareDirectoriesWindow menu
   - Share item/popup infrastructure with ContextMenu

6. **Eliminate ThemedControls.cpp menu hooks** — once menus are DxUI, remove all DRAWITEMSTRUCT/MEASUREITEMSTRUCT handlers

### Phase 3: Dialog Infrastructure (Medium Impact)
**Priority: P1 | Effort: 2-3 weeks**

7. **Build DxUI Toolbar** for RedSalamanderMonitor
   - Horizontal button strip with icons, separators, auto-size
   
8. **Add RadioButton** control (Toggle variant with group)
   - Needed for several Preferences pages

9. **Complete Preferences migration** — remaining pages still using HWND children
   - Keyboard.cpp and Themes.cpp have stale HWND gates (noted in Phase 2 DxUI review)
   - Use Grid for keyboard shortcut list, Tree for theme color list

### Phase 4: Viewer Plugin Modernization (Incremental)
**Priority: P1-P2 | Effort: 4-6 weeks | Per-plugin**

10. **Migrate viewer plugins to DxUI rendering surface**
    - Start with ViewerPE (simplest, 1 WndProc, 10 SB_ cases)
    - Then ViewerImgRaw (image + scrollbar)
    - Then ViewerText (3 WndProcs, most complex viewer)
    - ViewerVLC is special (5 WndProcs, VLC video surface) — keep native, wrap chrome only
    - ViewerSpace (D2D already, just menus + tooltip)

### Phase 5: Full Window Migration (Long-term)
**Priority: P2 | Effort: 6-12 months**

11. **Migrate CompareDirectoriesWindow** to full DxUI
    - Needs: Split container, ProgressBar, Grid, scroll, double-buffered rendering
    
12. **Migrate ConnectionManagerDialog** to full DxUI
    - Most complex dialog: card layout, ListView with 35+ operations, custom painting

13. **Migrate FolderWindow shell chrome** (StatusBar, layout, splitter)

14. **Evaluate NavigationView** — heavily custom (7 HWNDs, popups, edit suggestions)
    - May benefit from partial DxUI (suggestion popup, path display) while keeping edit as Win32

### Dependencies Graph
```
Phase 1 (Scrollbar, StatusBar, ProgressBar)
  └─→ Phase 4 (Viewer plugins need standalone scrollbar)
  └─→ Phase 3.9 (Preferences need all controls)
  
Phase 2 (Menu system)
  └─→ Phase 4 (Viewer plugins have custom menus)
  └─→ Phase 5 (All windows have menus)
  
Phase 3 (Toolbar, RadioButton)
  └─→ Phase 5 (Monitor needs toolbar before full migration)
```

---

## Key Metrics

| Metric | Count |
|--------|-------|
| DxUI controls today | 18 |
| New DxUI controls needed (P0+P1) | 6 (Scrollbar, Menu, ContextMenu, Toolbar, StatusBar sections, ProgressBar) |
| Win32 common control usages | 257 across 33 files |
| GDI drawing API calls | ~300+ across 25+ files |
| Custom WndProc implementations | 75+ |
| HWND member variables | 90+ |
| Lines of GDI menu drawing code | ~500 (single largest target) |
| Window class registrations | ~60+ |

---

## Decision

The DxUI framework is production-ready and covers the most important control types (Grid, Tree, ComboBox, TextField, Button/Checkbox). The **two critical gaps** are:

1. **Standalone Scrollbar** — blocks 8 files with 93 scrollbar message handlers
2. **Menu System** — blocks 15 files with ~500 lines of the most complex GDI code

Recommend starting Phase 1+2 in parallel: scrollbar extraction is mechanical and safe, while the menu system is a design challenge that benefits from early prototyping. Phase 3-5 are incremental and can proceed per-component as DxUI primitives become available.

**ThemedControls.cpp becomes fully dead code** once menus and the remaining Preferences pages migrate to DxUI — it's 2,800+ lines of Win32 theme hooks that serve no purpose in a DxUI world.
