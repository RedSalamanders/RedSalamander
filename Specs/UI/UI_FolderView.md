# FolderView Window Specification

## Overview

**FolderView** is a high-performance DirectX-based file browser component that displays folder contents in a grid layout with icons. It provides the primary file browsing interface for RedSalamander, supporting drag-and-drop, multi-selection, keyboard navigation, and context menus.

**Key Features:**
- Hardware-accelerated rendering (Direct2D, Direct3D 11, DXGI 1.3)
- Asynchronous folder enumeration and icon loading
- Grid layout with dynamic column sizing
- Sorting (Name / Extension / Time / Size / Attributes) with direction toggle + unsorted state
- Display modes: **Brief**, **Detailed**, **Extra Detailed**, and **Thumbnails**
- Pane filter (wildcard mask) with persisted per-path state and a subtle watermark indicator when active
- View options for hidden/system files (hidden items display a dim icon when shown)
- Thumbnail display mode with asynchronous shell/WIC thumbnail loading, aspect-preserving draw, icon fallback, and per-pane size persistence
- Full drag-and-drop support (COM IDataObject/IDropSource/IDropTarget)
- Multi-selection with visual feedback
- Keyboard navigation
- Per-monitor DPI awareness

**Architecture:**
- **Rendering**: D3D11 swap chain with Direct2D surface rendering
- **Startup performance**: D3D/D2D device + swap chain initialization is **deferred until after the first paint** (via `WndMsg::kFolderViewDeferredInit`) to keep `WM_CREATE` fast; first paint uses a GDI background fill until Direct2D is ready.
- **Threading**: Background enumeration thread for non-blocking folder loading
- **Icon Management**: Async icon loading **grouped by system icon index**; cached bitmaps are stamped immediately and missing icons are extracted once (background) + converted once (UI) then applied to all matching items. Icon bitmap conversion begins once a Direct2D device context is ready (no synchronous icon bitmap pre-warm during startup) and MUST NOT be blocked by pending swap-chain resize completion; otherwise fast startup can leave enumerated items stuck on placeholder icons.
- **Thumbnail Management**: Thumbnails is an exclusive display mode per pane. It keeps normal icon loading as fallback, queues only bounded visible work, extracts shell thumbnails off the UI thread, decodes likely image files through WIC when shell thumbnail extraction fails, creates Direct2D bitmaps on the UI thread from bounded pixel buffers, and drops stale work when navigation, refresh, sorting, display-mode, or thumbnail-size changes make a payload obsolete.
- **Parent-Child**: Child window of main application, coordinates with NavigationView

## Typography Contract

- FolderView labels/details must use the shared Windows 11 typography helper, not hardcoded `Segoe UI`.
- Visible text uses the shared Segoe UI Variable split:
  - details/caption-scale text may resolve to `Segoe UI Variable Small`
  - normal row labels and overlays use `Segoe UI Variable Text`
  - large overlay titles follow the shared size-based mapping
- Watermark and status glyphs use `Segoe Fluent Icons`.
- Any future visible HWND/GDI text added to FolderView-owned surfaces must route through the shared HFONT helper instead of `DEFAULT_GUI_FONT`.

## Performance Validation Contract

FolderView is a primary hot path and MUST be treated as performance-critical for new feature work.

Any new FolderView feature or optimization that can affect:

- cold or warm folder open,
- enumeration,
- sorting,
- icon loading,
- rendering,
- selection/scroll responsiveness,
- filter behavior,
- empty-state or overlay behavior,

MUST:

- define the protected scenario up front,
- add or reuse measurable instrumentation,
- add deterministic selftest coverage or another deterministic harness,
- archive validation runs under `Specs/TestRuns/`,
- use archived evidence when claiming improvement.

This requirement applies even when the first landing only establishes a baseline.

FolderView work SHOULD prefer existing metric families such as `render.*`, `icons.*`, and enumeration/sort metrics when they cover the scenario. If they do not, the feature MUST add the missing instrumentation with the change.

## Architecture

### Component Type
- **Class**: `FolderView`
- **Window Type**: Win32 child window with custom window class
- **Rendering**: Direct3D 11 + Direct2D 1.1 on DXGI swap chain
- **Parent**: FolderWindow (below the pane’s NavigationView)

### Files
- **Header**: `RedSalamander/FolderView.h`
- **Internal Helpers**: `RedSalamander/FolderViewInternal.h`
- **Helper Headers**:
  - `RedSalamander/FolderViewEmptyStateLayout.h`
  - `RedSalamander/FolderViewIncrementalSearch.h`
  - `RedSalamander/FolderViewVisualState.h`
- **Implementation (split)**:
  - `RedSalamander/FolderView.cpp` (window lifecycle + message dispatch)
  - `RedSalamander/FolderView.Interaction.cpp` (mouse/keyboard/scroll/command handling)
  - `RedSalamander/FolderView.Rendering.cpp` (swapchain/D2D/DWrite resources + rendering)
  - `RedSalamander/FolderView.Layout.cpp` (grid layout + hit testing + scroll metrics)
  - `RedSalamander/FolderView.Selection.cpp` (selection/focus + selection stats)
  - `RedSalamander/FolderView.Enumeration.cpp` (background enumeration + sorting + cache refresh)
  - `RedSalamander/FolderView.Icons.cpp` (async icon loading + UI-thread bitmap creation)
  - `RedSalamander/FolderView.Menus.cpp` (context menu + owner-draw menu theming)
  - `RedSalamander/FolderView.DragDrop.cpp` (drag source + drop target)
  - `RedSalamander/FolderView.FileOps.cpp` (delete/copy/paste/rename/move/properties)
  - `RedSalamander/FolderView.ErrorOverlay.cpp` (error reporting + overlay rendering)
- **Integration**: Created by FolderWindow, receives path updates from the paired NavigationView

### Component Interaction

``` console
┌──────────────────────────────────────────────────────────────────┐
│  Main Window                                                     │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │  NavigationView (24 DIP height)                             │ │
│  │  (Breadcrumb path display)                                  │ │
│  └─────────────────────────────────────────────────────────────┘ │
│  ┌─────────────────────────────────────────────────────────────┐ │ 
│  │  FolderView (fills remaining)                               │ │
│  │  ┌────────────────┐  ┌────────────────┐  ┌────────────────┐ │ │   
│  │  │📁 Pics        │  │📄 Text         │  │📄 Other Files │ │ │
│  │  │📁 Docs        │  │📄 Files        │  │📄 Other Files2│ │ │
│  │  │📄 File3       │  │📄 Files        │  │                │ │ │
│  │  └────────────────┘  └────────────────┘  └────────────────┘ │ │
│  └─────────────────────────────────────────────────────────────┘ │
└──────────────────────────────────────────────────────────────────┘
```

**Dual-pane host:**
- In dual-pane mode, `FolderWindow` hosts **two** independent pairs (Left/Right) of `NavigationView` + `FolderView` side-by-side.
- All interactions described below apply per `FolderView` instance (per pane).

**Callbacks:**
- NavigationView → FolderView: Path change via `SetFolderPath()`
- FolderView → NavigationView: Double-click activates the focused item:
  - Folder: navigate into it (updates `FolderView` path, which updates the paired `NavigationView`).
  - File: invoke the host’s open-file callback first (used to mount virtual file systems like `7z:`), otherwise fall back to `ShellExecuteW("open")`.

## Visual Layout and Grid System

### Grid Layout Algorithm

**Column-Based Arrangement:**
- Items arranged in vertical columns (top to bottom)
- Columns placed left to right
- Horizontal scrolling only (no vertical scroll)

**Column Width Calculation:**
```text
rowsPerColumn = floor((visiblePaneHeight + rowSpacing) / (itemHeight + rowSpacing))
items are assigned to columns top-to-bottom after rowsPerColumn is known

if (mode == Brief):
  textWidth[column] = max(nameWidth for each item assigned to that column)

if (mode == Detailed):
  // Ensure both lines fit (name + details), but only for the current column.
  nameWidth[column] = max(nameWidth for each item assigned to that column)
  detailsWidth[column] = max(detailsWidth for each item assigned to that column)
  textWidth[column] = max(nameWidth[column], detailsWidth[column])

if (mode == ExtraDetailed or mode == Thumbnails):
  // Ensure all three text lines fit (name + details + metadata), but only for the current column.
  nameWidth[column] = max(nameWidth for each item assigned to that column)
  detailsWidth[column] = max(detailsWidth for each item assigned to that column)
  metadataWidth[column] = max(metadataWidth for each item assigned to that column)
  textWidth[column] = max(nameWidth[column], detailsWidth[column], metadataWidth[column])

columnWidth[column] = max(textWidth[column] + iconWidth + padding, minColumnWidth)
columnWidth[column] = min(columnWidth[column], windowWidth)  // Don't exceed window width
```
- A wide item in one visible column MUST NOT widen earlier or later columns.
- Column left/right bounds are variable-width layout data and MUST be used by
  rendering, visible-range lookup, hit testing, horizontal scrolling, and
  `EnsureVisible`; these paths must not derive item positions from a single
  global column stride.
- Horizontal offset `0` is the canonical first-column scroll stop and preserves
  the leading column gutter. The first right line-scroll from `0` MUST move to
  the second column's left edge, not to the first column's left gutter; the
  matching first left line-scroll from the second column MUST return to `0` in
  one step.

**Item Spacing:**
- **Vertical spacing**:
  - Standard density: `4 DIP` between rows
  - Compact mode: `0 DIP` between rows so adjacent item bounds touch vertically
- **Horizontal spacing**: `18 DIP` between columns
- **Padding**: 8 DIP around icon and text
- Switching compact mode on an existing pane MUST relayout the visible grid immediately so rendering, hit testing, focus cues, and keyboard/mouse interactions all use the same row spacing in the new density.

**Text Truncation:**
- If filename exceeds column width, truncate with ellipsis ("...")
- Ellipsis rendered at end of visible text
- Tooltip shows full filename on hover (future enhancement)

### Visual Design

**Item Rendering (Brief):**
```text
    File name right of icon vertical center
              ↓
┌──────────────────────┐
│    🖼️ Filename.txt  │   
│   ICON               │ 
└──────────────────────┘
     ↑
16×16 DIP icon from shell

```

**Item Rendering (Thumbnail mode):**
```text
┌──────────────────────────────┐
│  [ 64 DIP thumbnail/icon ]   │
│  Filename.ext                │
└──────────────────────────────┘
```

- Thumbnail mode uses larger DPI-aware visuals for the item bitmap area. The per-pane size is one of `48`, `64`, `96`, or `128 DIP`, defaulting to `64 DIP`.
- Thumbnail bitmaps MUST preserve their source aspect ratio inside the thumbnail slot. Landscape images are letterboxed, portrait images are pillarboxed, and the draw rect MUST NOT exceed the slot.
- For local files, the pane requests shell thumbnails asynchronously for visible items only. If shell thumbnail extraction fails for a likely WIC-supported image extension, the worker decodes a bounded first-frame thumbnail through WIC before falling back to the normal icon.
- Until a thumbnail is ready, or when neither shell nor WIC can provide one, the normal folder/file icon is rendered in the same larger visual slot.
- Folder/file icon bitmap caching MUST keep separate entries for the shell image-list size class selected for the current target DIP size. Returning from thumbnail mode to Brief/Detailed/Extra Detailed MUST NOT reuse thumbnail-mode jumbo icon bitmaps in the normal `16 DIP` icon slot.
- Navigation, refresh, sorting, horizontal scrolling, ensure-visible focus moves, pane resize, thumbnail-size changes, and leaving thumbnail display mode cancel, invalidate, or requeue pending thumbnail work so stale payloads are not applied and newly visible columns do not remain on icon fallback.
- Thumbnail queues and caches are bounded to visible work and a memory budget; offscreen thumbnails may be evicted, but currently visible thumbnails must not be evicted during the same queue/paint pass.

**Item Rendering (Detailed):**
```text
┌──────────────────────────────────────────┐
│ 🖼️ Filename.txt                          │
│    TIME • (SIZE •) ATTRS                 │
└──────────────────────────────────────────┘
```

- Folder: `TIME • ATTRS`
- File: `TIME • SIZE • ATTRS`

**Item Rendering (Extra Detailed):**
```text
┌──────────────────────────────────────────┐
│ 🖼️ Filename.txt                          │
│    DETAILS (caller-provided)             │
│    METADATA (caller-provided)            │
└──────────────────────────────────────────┘
```

- Extra Detailed is a three-line layout intended for views that need both:
  - a primary “details” line (status/reason text), and
  - a secondary “metadata” line (time/size/attributes, etc.).
- The metadata line is optional; if no metadata provider is configured (or it returns empty), Extra Detailed behaves like Detailed.
- Thumbnails is also a three-line layout: name, details line, then metadata line. It uses the larger thumbnail/icon visual slot and may show metadata such as mount-point state, image dimensions, or other provider-specific information.

**Selection States:**
- **Normal**: Transparent background (theme-defined)
- **Hovered**: Light blue background (theme-defined)
- **Selected**: Accent color background (theme-defined)
- **Focused**: 2 DIP border (theme-defined); when the item is also **Selected**, the border uses a contrasting color (e.g., selected text color) to remain visible.
- Selection/hover backgrounds and the focus border use small rounded corners (see `Specs/UI/UI_VisualStyle.md`).
- **Unfocused pane**: normal, unselected item text/details/metadata and icons must render dimmer than the focused pane so the inactive FolderView is visually clear. Selected items use inactive-selection text/background colors instead of normal dimmed text, and the current item keeps a thinner/dimmer focus border.

**Multi-Selection Visual:**
- Multiple items show selection background
- Focused item has additional border
- Selection persists when focus moves to another item

## Features
### 1. Folder Content Display

**Data Source:**
- Uses plugin system (`IFileSystem::ReadDirectoryInfo`) **exclusively** for folder enumeration
- Direct enumeration via native APIs (including `std::filesystem::directory_iterator`, `FindFirstFileW`, etc.) is **prohibited** (no native fallback)
- If plugin is unavailable or enumeration fails, display a friendly **in-window alert overlay** (no `MessageBoxW`) using the shared `RedSalamander::Ui::AlertOverlay` component (`RedSalamander/Ui/AlertOverlay.h`):
  - Dimmed scrim + centered card, icon, title, and wrapped details text.
  - Pane-scoped modal (default): the current pane ignores normal interactions while the overlay is shown; the user can dismiss via the close “X” (errors/warnings/info).
- Alert overlays must **never disappear** due to small pane size; when space is constrained, the overlay switches to a **text-only** layout (icon hidden) and clips text as needed.
- If enumeration is slow (>300ms), show a **busy overlay** (spinner + “Please wait…” message), cleared when enumeration completes.
  - Busy overlays are not closable; for enumeration, the overlay exposes a **Cancel** action (button) to abort enumeration.
  - After the user cancels enumeration, show a non-dismissible **information** overlay (no close button) indicating the cancellation; it is cleared on the next navigation.
- Async enumeration on background thread
- **Empty-state message (non-error)**:
  - The host may set an optional empty-state message per pane (e.g., “This folder doesn’t exist in this hierarchy.”).
  - When enumeration succeeds and the visible item list is empty, and an empty-state message is set, FolderView renders that message centered in the client area using a secondary/dimmed text style.
  - The empty-state message must not replace error/busy overlays.
  - When enumeration succeeds and the visible item list is empty, and **no** host message is set, FolderView renders an **Empty folder** state:
    - A very dim, large watermark glyph (Segoe Fluent Icons `0xE8FF` "Preview").
    - Title text: **Empty folder**.
    - A fun, friendly message with a large emoji, chosen randomly from a small set of resource strings.
    - Double-click anywhere in the pane navigates **up to the parent folder**.
    - The focused empty-folder placeholder item is drawn as a normal item row at the top of the pane with localized label text **Go to parent**. Its focus/selection cue MUST span the current pane row width and use the current display mode's row height; it MUST NOT expand to the full pane body and MUST NOT shrink to a compact label-sized tile.
    - Empty-folder placeholder metrics MUST be recomputed from the current client width, DPI, icon size, and display-mode text-line heights. They MUST NOT inherit tile width/height from the previously displayed non-empty folder or from a previous Brief/Detailed/Extra Detailed/Thumbnails mode.

**View options & filtering:**
- **Hidden/System visibility** is controlled by settings:
  - `folders.showHiddenFiles` (default `true`)
  - `folders.showSystemFiles` (default `true`)
- When `folders.showHiddenFiles` is `false`, items with `FILE_ATTRIBUTE_HIDDEN` MUST be excluded from the item list.
- When `folders.showSystemFiles` is `false`, items with `FILE_ATTRIBUTE_SYSTEM` MUST be excluded from the item list.
- When hidden items are shown, they MUST display a **dim** icon to distinguish them from normal items.
- **Pane filter** (`cmd/pane/filter`): when enabled and the filter text parses to at least one mask, FolderView MUST exclude items whose `displayName` does not match the wildcard mask set (same syntax as Select/Unselect dialogs; see `MaskSyntax::MatchesWildcardMask`).
- **Hide names (in-memory)**:
  - `cmd/pane/selection/hideSelectedNames`: add the **currently selected** item `displayName`s to an in-memory hidden-names set (excludes those names from the view).
  - `cmd/pane/selection/hideUnselectedNames`: add the **currently unselected** item `displayName`s to the hidden-names set (no-op when nothing is selected).
  - `cmd/pane/selection/showHiddenNames`: clears the hidden-names set but MUST NOT change any existing pane filter.
  - The hidden-names set is pane-local and MUST be cleared on folder navigation.
- While a pane filter is active **or** the hidden-names set is non-empty, FolderView SHOULD render a very subtle watermark glyph in the background (Sego UI Symbol `0xE71C`) so the user can tell the pane is filtered/hidden.
- When the pane filter is active and no rows are visible, the filter state has priority over the generic Empty folder placeholder:
  - FolderView MUST suppress the generic Empty folder centered UI and its parent-row placeholder so the pane does not imply that the directory itself is empty.
  - FolderView MUST render the full-pane funnel watermark/placeholder for the filtered-empty result, unless an error/busy overlay is visible.
  - For explicit host empty-state messages that remain visible while filtering, FolderView MUST avoid overlapping watermarks and SHOULD use a small filter badge instead.

**Enumeration Contract (Plugin Only):**
- The host obtains an `IFileSystem` instance via the plugin factory (`RedSalamanderCreate(..., pluginId, ...)`) and uses it as the only source of directory entries.
- Each enumeration calls `IFileSystem::ReadDirectoryInfo(path, info.put())` to obtain an `IFilesInformation` result object.
- The returned `FileInfo` buffer is traversed via `NextEntryOffset` (preferred) to build `FolderItem` entries.

**Traversal Example (NextEntryOffset):**
```cpp
wil::com_ptr<IFilesInformation> info;
THROW_IF_FAILED(fileSystem->ReadDirectoryInfo(folder.c_str(), info.put()));

FileInfo* entry = nullptr;
THROW_IF_FAILED(info->GetBuffer(&entry));

while (entry != nullptr)
{
    // Build FolderItem from entry->FileName + entry->FileNameSize, entry->FileAttributes, etc.

    if (entry->NextEntryOffset == 0)
    {
        break;
    }

    entry = reinterpret_cast<FileInfo*>(
        reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset);
}
```

**Supported Item Types:**
- Files (all extensions)
- Folders/directories
- Network paths (UNC paths)
- Special folders (Desktop, Documents, etc.)

**Icon Rendering:**
- Uses IconCache system image lists (`SHIL_SMALL`/`SHIL_LARGE`/`SHIL_EXTRALARGE`/`SHIL_JUMBO`) and selects the **optimal** list size based on the target icon DIP size and current DPI (FolderView default is 16 DIP list-mode icons).
- IconCache MUST choose the smallest shell image-list source that is at least as large as the target physical pixel size; do not upscale 16px shell icons on high-DPI displays.
- FolderView icon bitmap draws MUST use nearest-neighbor interpolation only when the source bitmap and destination physical pixel size match exactly; scaled shell icons use linear filtering to avoid chunky edges.
- IconCache D2D bitmap normalization MUST premultiply translucent BGRA pixels and, for shell icons that carry RGB data but no alpha channel, apply Windows icon AND mask semantics: black mask pixels are opaque and non-black mask pixels are transparent.
- Fallback chain: **optimal → remaining sizes** (best-effort quality preservation)
- Icons cached in IconCache component (LRU cache, 2000 icons ≈18MB)
- Async loading with viewport prioritization: visible items first, offscreen queued
- Per-file icon extraction for .exe, .dll, .ico, .lnk, .url (embedded icons)
- Extension-based caching for common file types (bypasses Shell API on cache hit)
- Common startup warming includes the usual text/code types plus structured-log extensions such as `.jsonl` and `.ndjson`
- Fluent Design placeholder icons (folder: blue gradient, file: white document with fold)
- Shortcut overlay rendering for .lnk files (system SIID_LINK arrow)
- Performance telemetry: tracks cache hits, extraction count, load duration

**Performance:**
- Supports folders with 10,000+ items efficiently
- Virtualized rendering (only visible items laid out)
- Icon loading throttled to avoid UI freeze
### 2. Drag-and-Drop Support

**Implementation:** Full COM-based drag-and-drop using Windows Shell APIs

#### Drag Source (Dragging OUT of FolderView)

**COM Interfaces:**
- `FolderViewDataObject`: Implements `IDataObject` for clipboard formats
- `FolderViewDropSource`: Implements `IDropSource` for drag feedback

**Supported Clipboard Formats:**
- `CF_HDROP`: Shell-compatible file list (most important)
- `CFSTR_SHELLIDLIST`: Shell ID list for advanced operations
- `CFSTR_PREFERREDDROPEFFECT`: Suggests copy vs. move

**Drag Initiation:**
```cpp
// On mouse drag detected:
wil::com_ptr<IDataObject> dataObj = CreateDataObject(selectedItems);
wil::com_ptr<IDropSource> dropSource = CreateDropSource();
DWORD effect;
DoDragDrop(dataObj.get(), dropSource.get(), 
           DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK, &effect);
```

**Visual Feedback:**
- System-provided drag image (ghost icon + file count badge)
- Cursor changes based on drop target (copy/move/no-drop)

#### Drop Target (Dragging INTO FolderView)

**COM Interface:**
- `FolderView::DropTarget`: Implements `IDropTarget`

**Drop Operations:**
- **Copy**: Ctrl key held during drop
- **Move**: Default (no modifiers)
- **Link**: Ctrl+Shift keys held
- **Cancel**: Escape key

**Drop Validation:**
- Check if drop target is a folder
- Validate file system supports operation
- Highlight drop target folder during hover
- Show appropriate cursor (copy/move/no-drop arrow)
### 3. Context Menus

**Trigger:** Right-click on file/folder or selected items

**Localization requirement:** the context menu template is defined in `.rc` resources (see `Specs/Core/Core_Localization.md`) and loaded at runtime.

**Resource menu:**
- `IDR_FOLDERVIEW_CONTEXT` in `RedSalamander/RedSalamander.rc`

**Menu Items (v1):**
- Open
- Open With…
- Delete
- Move…
- Rename
- Copy
- Paste
- Properties

**Runtime behavior:**
- Items are enabled/disabled based on selection state (or `Current item` when selection is empty) and clipboard state.
- Menu rendering uses the active `MenuTheme` (owner-draw for themed background/selection colors).
### 4. Keyboard Navigation

**Canonical shortcut map**: `Specs/UI/UI_CommandMenuKeyboard.md` is the source of truth for global shortcuts and routing; this section focuses on FolderView-specific behavior.

**Arrow Key Navigation:**
- **Left/Right**: Move between columns
- **Up/Down**: Move within column
- **Home**: First item in first column
- **End**: Last item in last column
- **Page Up/Down**: Horizontal paging by **visible columns** (column layout)
- Moving the `Current item` does **not** change the selection set; focus and selection are independent.

**Pane Switching (FolderWindow Integration):**
- **Tab / Shift+Tab**: Switch focus to the **other pane**’s FolderView.
- Outside explicit child-window editing or popup flows, the active pane’s `FolderView` remains the default keyboard target for pane activation and application reactivation.

**NavigationView Access (FolderWindow Integration):**
- **F4 / Alt+D / Ctrl+L**: Focus the active pane’s NavigationView address bar and enter edit mode (select all). *(Default chord bindings are settings-backed.)*

**Selection Keys:**
- **Space**: Toggle selection of focused item
- **Insert**: Toggle selection of focused item and move to next item (Commander-style)
- **Alt+Up / Alt+Down**: Go to previous/next selected name (wrap allowed). *(Default chord binding is settings-backed.)*
- **Ctrl+Shift+<key left of Backspace>**: Select all items with the same extension as the focused item (adds to selection).
- **Ctrl+Shift+<key right of 0>**: Unselect all items with the same extension as the focused item (removes from selection).
- **Ctrl+A**: Select all items *(default chord binding is settings-backed)*
- **Ctrl+Click**: Toggle individual item selection
- **Shift+Click**: Range selection from anchor to clicked item
- **Ctrl+Shift+Arrow** or **Shift+Arrow**: Extend selection without moving focus
- **Ctrl+Shift+Click** : Extend selection without moving focus
- **Shift+Home/End**: Extend selection to start/end
- **Esc**: Clear selection


**Action Keys:**
- **Enter**: Open focused item (folder navigates; file invokes host open hook which may mount a virtual file system or fall back to `ShellExecute`)
- **Delete**: Delete selected items (with confirmation)
- **F2**: Rename focused item
- **Backspace**: Navigate to parent folder
- **F3/F5/F6/F7/F8**: FolderWindow-global operations (view/copy/move/mkdir/delete) per `Specs/UI/UI_CommandMenuKeyboard.md`.

**Clipboard Keys:**
- **Ctrl+C**: Copy selected items to clipboard (or `Current item` when selection is empty) *(default chord binding is settings-backed)*
- **Ctrl+V**: Paste from clipboard to current folder *(default chord binding is settings-backed)*

**View/Sort Commands (FolderWindow Integration):**
- **Ctrl+F2**: Sort by **None** (restore initial order)
- **Ctrl+F3**: Sort by **Name**
- **Ctrl+F4**: Sort by **Extension**
- **Ctrl+F5**: Sort by **Time** (newest first)
- **Ctrl+F6**: Sort by **Size** (largest first; folders fall back to Name)
- **Ctrl+F12**: Open the pane filter dialog (wildcard mask filter; affects enumeration).
- **Alt+2**: Display as **Brief**
- **Alt+3**: Display as **Detailed**
- **Alt+4**: Display as **Extra Detailed**
- **Alt+5**: Display as **Thumbnails**
- Sort by **Attributes** is currently menu-only (no default shortcut).

**Notes:**
- In dual-pane mode, these commands apply to the **focused** pane.
- Sorting is **directories-first**, then by the selected key.
- Reselecting the same sort key toggles direction: default direction ↔ opposite direction (use **None** / `Ctrl+F2` to restore the initial order).

### Incremental Search (FolderView)

FolderView implements **incremental search mode** (type-to-search) as specified in `Specs/UI/UI_CommandMenuKeyboard.md`:
- Typing printable characters updates the search query. Highlighting uses a case-insensitive substring match across all visible item display names, but focus navigation uses a case-insensitive prefix match: the current item jumps to the next visible item whose name starts with the query.
- Matching text is highlighted while the mode is active.
  - All **visible** items whose display name matches the query show the highlight on the matched substring.
  - Highlight style: the matched substring gets a **selection-style background** (use `itemBackgroundSelected` / `itemBackgroundSelectedInactive`) with the corresponding selection text color; do **not** change font weight.
  - If the item is **selected**, the highlight still renders (use a subtle contrasting in-selection background scrim) and must not override the selected text color.
- Backspace edits the query; Esc exits the mode.

### 5. Multi-Selection

**Selection Modes:**
1. **Single-click**: Move focus to clicked item (does not change selection)
2. **Ctrl+Click**: Toggle clicked item selection, keep others
3. **Shift+Click**: Select range from anchor to clicked item (creates selection)
4. **Marquee**: Drag on empty space to select multiple (future)

**Visual Feedback:**
- On folder entry (after enumeration), FolderView sets the `Current item` to the first item (or nearest preserved focus), and starts with **no selection**.
- On refresh of the currently displayed folder, including refreshes triggered by `DirectoryInfoCache` callbacks, the user selection is preserved for every surviving item whose display name is still present even when size, time, attributes, details, layout, or icon rendering state changed. Items that disappeared from the refreshed list are removed from the selection. When the refresh impact carries chained same-folder rename hints, a selected old display name MUST transfer selection through the full chain to the final new display name; an unselected original MUST NOT become selected merely because it was renamed. When no rename hint is available, the old name is treated as removed.
- Selected items: Selection background differs between the focused vs unfocused pane (subtle inactive selection), per `Specs/UI/UI_CommandMenuKeyboard.md`.
- Current item: Focus border always; in the focused pane it also has a background fill.
- Selected + Current item: Draw selection background plus a contrasting focus border stroke so the focus state remains visible on top of selection.
- Item state matrix (visual):
  - Normal: no fill, `textNormal`.
  - Selected (focused pane): `itemBackgroundSelected` + `textSelected`.
  - Selected (unfocused pane): `itemBackgroundSelectedInactive` + `textSelectedInactive`.
  - Focused (focused pane): focus border + `itemBackgroundFocused` fill.
  - Focused (unfocused pane): focus border only (thinner stroke + reduced opacity compared to the focused pane).
- Selected + Focused (focused pane): `itemBackgroundSelected` + focus border (contrasting).
- Selected + Focused (unfocused pane): `itemBackgroundSelectedInactive` + focus border (contrasting).
- Status bar (per pane): selection summary (folders/files + total selected bytes; folder sizes may be unknown until explicitly requested, then computed via an iterative folder-subtree traversal and can be “calculating”).
  - Example: `3 files: 4.60 MB selected`
  - Example: `2 folders / 5 files: 8.90 KB selected`

**Selection API:**
```cpp
void SelectItem(size_t index, bool clearOthers = true);
void ToggleSelection(size_t index);
void SelectRange(size_t start, size_t end);
void SelectAll();
void ClearSelection();
std::vector<size_t> GetSelectedIndices() const;
```        
## Theme System

### Color Theme Structure

All visual colors must be defined through a theme structure to support customization and the required built-in themes: **Light**, **Dark**, **Rainbow**, and **System High Contrast**.

**Theme Definition:**
```cpp
struct FolderViewTheme {
	    // Background colors
	    D2D1::ColorF backgroundColor;           // Main background (default: white)
	    D2D1::ColorF itemBackgroundNormal;      // Normal item background (transparent)
	    D2D1::ColorF itemBackgroundHovered;     // Hovered item background
	    D2D1::ColorF itemBackgroundSelected;    // Selected item background
	    D2D1::ColorF itemBackgroundSelectedInactive; // Selected item background in the unfocused pane
	    D2D1::ColorF itemBackgroundFocused;     // Focused item additional highlight
	    
	    // Text colors
	    D2D1::ColorF textNormal;                // Normal text color
	    D2D1::ColorF textSelected;              // Selected item text
	    D2D1::ColorF textSelectedInactive;      // Selected item text in the unfocused pane
	    D2D1::ColorF textDisabled;              // Disabled/unavailable items
    
    // Border and outline colors
    D2D1::ColorF focusBorder;               // Focus rectangle border
    D2D1::ColorF gridLines;                 // Grid/separator lines (if applicable)
    
    // Alert colors (overlay/banner)
    D2D1::ColorF errorBackground;           // Error message background
    D2D1::ColorF errorText;                 // Error message text
    D2D1::ColorF warningBackground;         // Warning message background
    D2D1::ColorF warningText;               // Warning message text
    D2D1::ColorF infoBackground;            // Info message background
    D2D1::ColorF infoText;                  // Info message text
    
    // Drag-and-drop feedback
    D2D1::ColorF dropTargetHighlight;       // Drop target folder highlight
    D2D1::ColorF dragSourceGhost;           // Dragged item ghost overlay
};
```

**Built-in Themes:**

- `itemBackgroundSelectedInactive` should be a subtle “inactive selection” variant of `itemBackgroundSelected` (recommended: same RGB with a reduced alpha, e.g. `0.65f`).
- `textSelectedInactive` must remain readable over the **composited** inactive selection background (recommended: pick a contrasting text color for the effective background).
- In the unfocused pane, the focus border is also rendered dimmer (recommended: multiply alpha by ~`0.60f`) in addition to the thinner stroke.

**Light (Default):**
```cpp
static FolderViewTheme GetDefaultLightTheme() {
    return FolderViewTheme{
        .backgroundColor = D2D1::ColorF(D2D1::ColorF::White),
        .itemBackgroundNormal = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f),  // Transparent
        .itemBackgroundHovered = D2D1::ColorF(0.902f, 0.941f, 1.0f),   // RGB(230, 240, 255)
        .itemBackgroundSelected = GetSystemAccentColor(),               // Windows accent color
        .itemBackgroundFocused = D2D1::ColorF(0.0f, 0.478f, 1.0f, 0.3f), // Semi-transparent blue
        
        .textNormal = D2D1::ColorF(D2D1::ColorF::Black),
        .textSelected = D2D1::ColorF(D2D1::ColorF::White),
        .textDisabled = D2D1::ColorF(0.6f, 0.6f, 0.6f),
        
        .focusBorder = GetSystemAccentColor(),
        .gridLines = D2D1::ColorF(0.9f, 0.9f, 0.9f),
        
        .errorBackground = D2D1::ColorF(1.0f, 0.95f, 0.95f),
        .errorText = D2D1::ColorF(0.8f, 0.0f, 0.0f),

        .warningBackground = D2D1::ColorF(1.0f, 0.98f, 0.90f),
        .warningText = D2D1::ColorF(0.65f, 0.38f, 0.0f),

        .infoBackground = D2D1::ColorF(0.90f, 0.95f, 1.0f),
        .infoText = GetSystemAccentColor(),
        
        .dropTargetHighlight = D2D1::ColorF(0.0f, 0.478f, 1.0f, 0.4f),
        .dragSourceGhost = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.5f)
    };
}

static D2D1::ColorF GetSystemAccentColor() {
    DWORD accentColor = 0;
    DWORD colorType = 0;
    SystemParametersInfoW(SPI_GETCOLORACCENTCOLOR, 0, &accentColor, 0);
    
    return D2D1::ColorF(
        ((accentColor >> 16) & 0xFF) / 255.0f,  // R
        ((accentColor >> 8) & 0xFF) / 255.0f,   // G
        (accentColor & 0xFF) / 255.0f,          // B
        1.0f
    );
}
```

**Dark:**
```cpp
static FolderViewTheme GetDefaultDarkTheme() {
    return FolderViewTheme{
        .backgroundColor = D2D1::ColorF(0.08f, 0.08f, 0.08f),
        .itemBackgroundNormal = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f),
        .itemBackgroundHovered = D2D1::ColorF(0.16f, 0.16f, 0.16f),
        .itemBackgroundSelected = GetSystemAccentColor(),
        .itemBackgroundFocused = D2D1::ColorF(0.0f, 0.478f, 1.0f, 0.25f),

        .textNormal = D2D1::ColorF(0.92f, 0.92f, 0.92f),
        .textSelected = D2D1::ColorF(D2D1::ColorF::White),
        .textDisabled = D2D1::ColorF(0.55f, 0.55f, 0.55f),

        .focusBorder = GetSystemAccentColor(),
        .gridLines = D2D1::ColorF(0.18f, 0.18f, 0.18f),

        .errorBackground = D2D1::ColorF(0.30f, 0.10f, 0.10f),
        .errorText = D2D1::ColorF(1.0f, 0.65f, 0.65f),

        .warningBackground = D2D1::ColorF(0.28f, 0.22f, 0.12f),
        .warningText = D2D1::ColorF(1.0f, 0.80f, 0.35f),

        .infoBackground = D2D1::ColorF(0.12f, 0.18f, 0.28f),
        .infoText = GetSystemAccentColor(),

        .dropTargetHighlight = D2D1::ColorF(0.0f, 0.478f, 1.0f, 0.35f),
        .dragSourceGhost = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.30f)
    };
}
```

**Rainbow:**
- Uses a neutral readable base (light or dark) and derives **accent/selection** colors from a hue cycle (e.g., item index or a stable hash of the full path).
- Must preserve text contrast; do not use rainbow colors for primary text.

**Theme Usage:**
```cpp
class FolderView {
private:
    FolderViewTheme _theme;
    wil::com_ptr<ID2D1SolidColorBrush> _backgroundBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _textBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _selectionBrush;
    // ... other brushes
    
public:
    void SetTheme(const FolderViewTheme& theme) {
        _theme = theme;
        RecreateThemeBrushes();
        InvalidateRect(_hWnd, nullptr, FALSE);
    }
    
private:
    void RecreateThemeBrushes() {
        _d2dDeviceContext->CreateSolidColorBrush(_theme.backgroundColor, &_backgroundBrush);
        _d2dDeviceContext->CreateSolidColorBrush(_theme.textNormal, &_textBrush);
        _d2dDeviceContext->CreateSolidColorBrush(_theme.itemBackgroundSelected, &_selectionBrush);
        // ... create all theme-dependent brushes
    }
    
    void RenderItem(const FolderItem& item, const D2D1_RECT_F& rect) {
        // Determine item state
        bool isSelected = IsItemSelected(item.index);
        bool isHovered = (_hoveredIndex == item.index);
        bool isFocused = (_focusedIndex == item.index);
        
        // Select appropriate colors from theme
        D2D1::ColorF bgColor = _theme.itemBackgroundNormal;
        D2D1::ColorF textColor = _theme.textNormal;
        
        if (isSelected) {
            bgColor = _theme.itemBackgroundSelected;
            textColor = _theme.textSelected;
        } else if (isHovered) {
            bgColor = _theme.itemBackgroundHovered;
        }
        
        // Render with theme colors
        wil::com_ptr<ID2D1SolidColorBrush> bgBrush;
        _d2dDeviceContext->CreateSolidColorBrush(bgColor, &bgBrush);
        _d2dDeviceContext->FillRectangle(rect, bgBrush.get());
        
        if (isFocused) {
            wil::com_ptr<ID2D1SolidColorBrush> focusBrush;
            _d2dDeviceContext->CreateSolidColorBrush(_theme.focusBorder, &focusBrush);
            _d2dDeviceContext->DrawRectangle(rect, focusBrush.get(), 2.0f);
        }
        
        // ... render icon and text with textColor
    }
};
```

### Theme Integration Points

**System Integration:**
- Monitor `WM_SETTINGCHANGE` / `WM_THEMECHANGED` for system theme and accent color changes
- Detect and apply Windows High Contrast mode (`SystemParametersInfoW(SPI_GETHIGHCONTRAST, ...)`) and override other themes when enabled
- Support Windows light/dark mode detection and/or explicit user theme selection (Light/Dark/Rainbow)

**Application Integration (Current Implementation):**
- Theme is selected at the application level via the top menu: **View → Theme** (radio-check marks).
- The menu includes built-in themes and any `user/*` themes found in settings and/or `Themes\\*.theme.json5` next to the executable.
- **High Contrast** is system-controlled and always overrides the selected theme:
  - Menu shows **High Contrast (System)** as **checked + disabled** when enabled.
  - Switching System/Light/Dark/Rainbow updates the radio check, but the effective palette remains High Contrast until Windows High Contrast is turned off.
- Theme changes are propagated through `FolderWindow::ApplyTheme()` to both FolderView and NavigationView, and the app titlebar is updated (best-effort).

**User Customization:**
- Expose theme through settings/configuration
- Allow custom color overrides per theme element
- Save/load theme preferences

**DPI Awareness:**
- Theme colors are DPI-independent (use normalized 0.0-1.0 values)
- Border widths and sizes scale with DPI
- Brushes recreated on DPI change maintain theme colors

## Placeholder Icons (Fluent Design)

**Purpose:** Show modern, high-quality placeholders while icons are loading asynchronously.

**Implementation:** Direct2D path geometry with explicit pixel format.

### Folder Placeholder
```cpp
// Blue gradient folder icon (48×48)
// Top: RGB(80, 148, 232) → Bottom: RGB(52, 120, 200)
void CreatePlaceholderIcon() {
    // Create bitmap render target with explicit pixel format
    D2D1::PixelFormat pixelFormat = D2D1::PixelFormat(
        DXGI_FORMAT_B8G8R8A8_UNORM,
        D2D1_ALPHA_MODE_PREMULTIPLIED);
    
    wil::com_ptr<ID2D1BitmapRenderTarget> folderTarget;
    _d2dContext->CreateCompatibleRenderTarget(
        &targetSize, nullptr, &pixelFormat, 
        D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_NONE, &folderTarget);
    
    // Draw rounded rectangle with gradient fill
    folderTarget->BeginDraw();
    folderTarget->Clear(D2D1::ColorF(0, 0, 0, 0)); // Transparent
    
    wil::com_ptr<ID2D1LinearGradientBrush> gradientBrush;
    // ... gradient setup ...
    
    folderTarget->FillRoundedRectangle(
        D2D1::RoundedRect(rect, 2.0f, 2.0f), 
        gradientBrush.get());
    
    folderTarget->EndDraw();
    folderTarget->GetBitmap(_placeholderFolderIcon.addressof());
}
```

### File Placeholder
```cpp
// White document with gray fold (48×48)
void CreatePlaceholderIcon() {
    // Create path geometry for document shape
    wil::com_ptr<ID2D1PathGeometry> docPath;
    _d2dFactory->CreatePathGeometry(&docPath);
    
    wil::com_ptr<ID2D1GeometrySink> sink;
    docPath->Open(&sink);
    
    sink->BeginFigure(D2D1::Point2F(12, 8), D2D1_FIGURE_BEGIN_FILLED);
    // ... path commands for document shape with corner fold ...
    sink->EndFigure(D2D1_FIGURE_END_CLOSED);
    sink->Close();
    
    // Fill with white, outline with gray
    fileTarget->FillGeometry(docPath.get(), fillBrush.get());
    fileTarget->DrawGeometry(docPath.get(), outlineBrush.get(), 1.0f);
    
    // Draw fold line
    fileTarget->DrawLine(
        D2D1::Point2F(30, 8), D2D1::Point2F(38, 16),
        outlineBrush.get(), 1.0f);
}
```

### Shortcut Overlay
```cpp
// Extract system shortcut arrow (16×16)
void CreatePlaceholderIcon() {
    SHSTOCKICONINFO sii{};
    sii.cbSize = sizeof(sii);
    HRESULT hr = SHGetStockIconInfo(SIID_LINK, SHGFI_ICON | SHGSI_SMALLICON, &sii);
    
    if (SUCCEEDED(hr) && sii.hIcon) {
        // Convert HICON to bitmap and cache
        // Rendered on top of file/folder icons for .lnk files
    }
}
```

**Critical:** Must specify explicit pixel format when creating render targets. Using `nullptr` for pixel format causes D2D debug layer error ("DXGI_FORMAT_UNKNOWN is not allowed").

**Rendering:** Placeholders shown immediately, replaced with actual icon when loaded. Shortcut overlay composited on top for .lnk files.

## Implementation Details

### DirectX Rendering Pipeline

**Device and Swap Chain Creation:**
```cpp
// D3D11 device
D3D11CreateDevice(..., D3D_DRIVER_TYPE_HARDWARE, ..., D3D_FEATURE_LEVEL_11_0, ...);

// DXGI swap chain
DXGI_SWAP_CHAIN_DESC1 desc = {};
desc.Format = DXGI_FORMAT_B8G8R8A8_UNORM;
desc.SwapEffect = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
desc.Flags = DXGI_SWAP_CHAIN_FLAG_ALLOW_TEARING;  // Variable refresh rate
dxgiFactory->CreateSwapChainForHwnd(_d3dDevice.get(), _hWnd, &desc, ...);

// D2D render target from swap chain
_d2dDeviceContext->CreateBitmapFromDxgiSurface(dxgiBackBuffer, &props, &_d2dTargetBitmap);
```

**Rendering Loop:**
```cpp
void OnPaint() {
    _d2dDeviceContext->BeginDraw();
    _d2dDeviceContext->Clear(_theme.backgroundColor);  // Use theme color
    
    // Render visible items only
    for (auto& item : GetVisibleItems()) {
        RenderItem(item);
    }
    
    _d2dDeviceContext->EndDraw();
    
    // Present with vsync (1) or immediate (0)
    _swapChain->Present(1, 0);
}
```

**DPI Handling:**
```cpp
void OnDpiChanged(float newDpi) {
    _currentDpi = newDpi;
    
    // Recreate DPI-dependent resources
    RecreateTextFormats(newDpi);
    RecreateIconBitmaps(newDpi);
    RecreateThemeBrushes();  // Recreate with same theme colors
    
    // Recalculate layout
    InvalidateLayout();
    InvalidateRect(_hWnd, nullptr, FALSE);
}
```

### Threading Model

**UI Thread:**
- Window message handling (WM_PAINT, WM_SIZE, etc.)
- DirectX rendering
- User input processing
- Icon bitmap creation

**Background Thread (Enumeration):**
- Folder enumeration via `IFileSystem::ReadDirectoryInfo`
- Posts `WndMsg::kFolderViewEnumerateComplete` when finished
- Payload contains file list

**Background Thread (Icon Loading):**
- Extracts `wil::unique_hicon` from the system image list via `IconCache::ExtractSystemIcon(...)` using the already-resolved `iconIndex` (**requires COM initialized as MTA on the worker thread**)
- Posts `WndMsg::kFolderViewCreateIconBitmap` with an `IconBitmapRequest` that owns a `wil::unique_hicon`
- UI thread converts `wil::unique_hicon` → `ID2D1Bitmap1` via `IconCache::ConvertIconToBitmapOnUIThread(hIcon.get(), ...)` and caches the result
- If the bitmap is already cached for the current D2D device, posts `kMsgIconLoaded` (UI thread just fetches from cache)

**Thread Synchronization:**
```cpp
// Enumeration thread
void EnumerateFolder(const std::filesystem::path& path) {
    auto payload = std::make_unique<EnumerationPayload>();
    payload->items = LoadFolderContentsViaPlugin(path); // Calls IFileSystem::ReadDirectoryInfo (no native fallback)

    // `PostMessagePayload` reclaims the payload automatically if PostMessageW fails.
    static_cast<void>(PostMessagePayload(_hWnd.get(), WndMsg::kFolderViewEnumerateComplete, 0, std::move(payload)));
}

// UI thread message handler
case WndMsg::kFolderViewEnumerateComplete: {
    auto payload = TakeMessagePayload<EnumerationPayload>(lParam);
    if (! payload) { return 0; }
    UpdateItemList(payload->items);
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
    return 0;
}
```

**Teardown note:** windows that receive payload messages should call `InitPostedPayloadWindow(hwnd)` on create and `DrainPostedPayloadsForWindow(hwnd)` in `WM_NCDESTROY` to prevent leak-on-destroy.

### IconCache Integration

**Architecture:**
- **Singleton instance** shared across all FolderView windows
- **LRU cache** of D2D bitmaps (2000 icons; worst-case ≈18MB at 48×48 BGRA, smaller at 16×16 or 32×32 depending on target DIP size/DPI)
- **Extension-to-iconIndex mapping** for instant lookups (bypasses Shell API)
- **Per-file whitelist** for unique icons (.exe, .dll, .ico, .lnk, .url, .cpl, .scr, .msc, .ocx)
- **3-level fallback** chain: selected optimal size → other sizes (from `SHIL_SMALL`/`SHIL_LARGE`/`SHIL_EXTRALARGE`)
- Extension-association misses fall back to the generic shell file/folder icon index instead of leaving `iconIndex = -1`.
- Transient `ExtractSystemIcon(...)` failures are retried before FolderView gives up and leaves the placeholder visible.

**Icon Extraction Flow:**
```cpp
// 1. Cached by extension (and auto-populates cache on demand)
auto iconIndex = IconCache::GetInstance().GetOrQueryIconIndexByExtension(extension, fileAttributes);
if (iconIndex.has_value())
{
    item.iconIndex = *iconIndex;
}

// 2. Per-file lookup types (.exe/.lnk/etc.) or special folders
else if (IconCache::GetInstance().RequiresPerFileLookup(extension) ||
         (item.isDirectory && IconCache::IsSpecialFolder(item.fullPath.wstring())))
{
    const std::wstring fullPath = item.fullPath.wstring();
    item.iconIndex = IconCache::GetInstance().QuerySysIconIndexForPath(fullPath.c_str(), 0, false).value_or(-1);
}

// 3. Convert icon index to D2D bitmap on UI thread (cached per D2D device)
auto bitmap = IconCache::GetInstance().GetIconBitmap(item.iconIndex, _d2dContext);
```

**Viewport-Aware Loading:**
- Visible items processed first (high priority)
- Offscreen items queued (low priority)
- Parallel extraction using Windows Thread Pool
- Telemetry logged: total requests, visible requests, cache hits, extracted count, duration. IconCache lock diagnostics MUST be thresholded slow-path rows (`iconcache.lock_wait_slow_us` / `iconcache.lock_hold_slow_us`) with the IconCache stage in `detail`; raw per-lock wait/hold rows are too noisy for archived perf runs and MUST NOT be emitted. Shell icon lookup timings (`iconcache.shgetfileinfo_us`) MUST identify the lookup kind in `detail` (`association`, `path_attributes`, or `path_live`) and carry file attributes plus SHGFI flags in the value fields.

**Cache Warming:**
- Common extensions pre-cached at startup (50+ types including .txt, .pdf, .zip, .jpg, etc.)
- Special folder icons pre-loaded (Desktop, Documents, Pictures, etc.)
- Reduces Shell API calls during folder enumeration

**Memory Management:**
- LRU eviction when cache exceeds 2000 icons
- Eviction triggers on every GetIconBitmap() call
- Memory footprint: ~9KB per 48×48 BGRA bitmap
- Total cache size: ≈18MB (2000 icons × 9KB)

**Performance Metrics:**
- Cache hit: ~1-5 microseconds (map lookup)
- Extension query: ~50-200 microseconds (SHGetFileInfo with SHGFI_USEFILEATTRIBUTES)
- Per-file extraction: ~1-5 milliseconds (reads file metadata)
- Parallel extraction: N files concurrently via Thread Pool

### Resource Management (RAII)

**WIL Smart Pointers:**
```cpp
wil::com_ptr<ID3D11Device> _d3dDevice;
wil::com_ptr<ID2D1DeviceContext> _d2dDeviceContext;
wil::com_ptr<IDXGISwapChain1> _swapChain;
wil::com_ptr<ID2D1Bitmap1> _d2dTargetBitmap;
wil::com_ptr<IDWriteTextFormat> _textFormat;

// Automatically released in destructor
```

**Device Loss Handling:**
```cpp
void HandleDeviceLost() {
    // Release all device-dependent resources
    _d2dTargetBitmap.reset();
    _d2dDeviceContext.reset();
    _swapChain.reset();
    
    // Recreate device and resources
    CreateDeviceResources();
    CreateWindowSizeDependentResources();
    
    InvalidateRect(_hWnd, nullptr, FALSE);
}
```

## Performance Optimizations

### Virtualization

**Only render visible items:**
```cpp
std::vector<FolderItem*> GetVisibleItems() {
    std::vector<FolderItem*> visible;
    
    // Calculate visible range based on scroll position
    int firstColumn = _scrollX / _columnWidth;
    int lastColumn = (_scrollX + _clientWidth) / _columnWidth + 1;
    
    // Return only items in visible columns
    for (int col = firstColumn; col <= lastColumn && col < _columnCount; col++) {
        for (auto& item : GetItemsInColumn(col)) {
            visible.push_back(&item);
        }
    }
    
    return visible;
}
```

### Dirty Region Tracking

**Invalidate only changed regions:**
```cpp
void OnItemSelectionChanged(size_t index) {
    RECT itemRect = GetItemRect(index);
    InvalidateRect(_hWnd, &itemRect, FALSE);  // Only redraw this item
}
```

### Batch Icon Loading

**Load icons in batches to avoid thread thrashing:**
```cpp
void RequestIconBatch(const std::vector<size_t>& indices) {
    constexpr size_t kBatchSize = 50;
    
    for (size_t i = 0; i < indices.size(); i += kBatchSize) {
        auto batch = std::vector<size_t>(
            indices.begin() + i,
            indices.begin() + std::min(i + kBatchSize, indices.size())
        );
        QueueIconLoadBatch(batch);
    }
}
```

## Error Handling

**File System Errors:**
- Access denied: Show in-window error overlay
- Path not found: Show in-window error overlay
- Disconnected/unavailable location (USB removed, mapped drive removed, network share unavailable): show an in-window **information** overlay (“Disconnected”) scoped to the FolderView that is **non-dismissible** and **blocks FolderView input** (navigation bar remains usable for recovery).
- Slow/unresponsive paths: Show busy overlay while enumerating; on failure, show error overlay

**Debug / Debugging Sessions:**
- When running in debug mode (Debug build), Left/Right menus include an “Overlay Sample” submenu to preview Error/Warning/Information/Busy overlays and hide them.
- If the menu bar is hidden, press `Alt` (or enable `View -> Menu Bar`) to access the Left/Right menus; the same sample actions are also available in the FolderView context menu.

**Rendering Errors:**
- Device lost: Recreate device and retry
- Out of memory: Reduce icon cache size, show error
- Swap chain recreation failure: Show error overlay and retry recreation (no GDI fallback; no `MessageBoxW`)

**Pattern:** report failures via the in-window overlay system (no `MessageBox*`).

Example: `FolderView::ReportError(L"EnumerateFolder", hr)` logs the failure and updates the current in-window overlay state.
## Testing

### Unit Tests
- Grid layout calculation with various window sizes
- Compact-mode density toggle collapses row spacing to `0 DIP` and updates hit testing immediately
- Selection state management (single, multi, range)
- Keyboard navigation logic
- Incremental search keeps contains-based highlights but uses prefix matching for focus/jump.
- Inactive-pane visual-state helpers dim normal text/icons while preserving selected/focused inactive states.
- Empty-folder placeholder metric helpers recompute placeholder item width/height for the current empty layout and clamp width to the current client area.
- Icon cache hit/miss rates

### Integration Tests
- Folder enumeration with plugin system
- Drag-and-drop between FolderView instances
- Context menu invocation
- DPI change handling

### Performance Tests
- Large folder (10K+ files) render time: <500ms
- Scroll smoothness: 60fps sustained
- Icon loading: 100 icons/second
- `folderView_perf_large_folder_baseline` must verify that stock icon loading resolves into item bitmap icons, not only that icon-index lookup or cache warming occurred.
- Memory usage: <100MB for 10K items

### Manual Tests
- Unicode filename display (Chinese, Arabic, emoji)
- Network path browsing (UNC paths)
- Removable media (USB drives)
- Special folders (Desktop, Documents)
- Shell extension context menus (7-Zip, Git)
## AGENTS.md Compliance

- **C++23**: Use `std::filesystem::path`, structured bindings, `constexpr`
- **RAII**: All resources managed via WIL and STL containers
- **Smart Pointers**: `wil::com_ptr` for COM, `std::unique_ptr` for ownership
- **No raw new/delete**: Use `std::make_unique` and containers
- **WIL Usage**: COM objects, Windows handles, error handling
- **Unicode UTF-16**: All strings are `std::wstring` or `wchar_t*`
- **Error Handling**: `THROW_IF_FAILED` for HRESULT errors
- **Threading**: Clear thread ownership, message-based async communication
- **Performance**: Virtualized rendering, dirty region tracking, async operations

## Future Enhancements

1. **List/Details View**: Table view with columns (Name, Size, Modified)
2. **Grouping**: Group by type, date, size
3. **Marquee Selection**: Drag rectangle to select multiple items
4. **Inline Rename**: Rename without dialog
5. **Quick Look**: Space bar to preview file without opening
6. **Column Resizing**: Drag column dividers to resize
7. **Multi-key Sorting**: Primary + secondary key (e.g. Type then Name)
8. **Pane-to-Pane Operations**: Copy/move between panes
