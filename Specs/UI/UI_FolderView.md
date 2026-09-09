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
- FolderView-owned visible text must not introduce an HWND/GDI/HFONT fallback. It stays on the shared DxUi/DirectWrite typography path; any required native interop must be reviewed through `UI_VisibleTypographyAudit.md` and `UI_VisibleNativeAudit.md`.

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

FolderView frame instrumentation is now part of that contract. Rendering paths must emit bounded per-frame aggregates for `folder.frame.total_us`, `folder.frame.present_us`, `folder.frame.visible_work_count`, `folder.frame.input_to_paint_us`, and `folder.frame.dirty_rect_area_px` when the scenario exercises a real FolderView frame. Overlay paint paths must additionally expose `folder.frame.overlay_animation_count` and `folder.frame.overlay_dirty_rect_area_px` when overlay animation or incremental-search overlay work actually renders. The older `render.*` rows remain useful for detailed renderer stages, including `render.incremental_search_effect_updates`, but claims about user-visible frame responsiveness should cite the `folder.frame.*` and input-to-paint metrics. FolderView input-to-paint and refresh-to-paint pending metrics may share a single storage shape, but they MUST remain separate pending slots so refresh telemetry cannot overwrite or consume navigation/input latency telemetry.

Settled overlay animation ticks must use bounded invalidation. After the initial show/scrim transition settles, busy/error overlay animation should invalidate the union of the previous and current overlay panel bounds plus the active incremental-search badge bounds, not the full FolderView client. Full-client invalidation remains valid for show/dismiss transitions, full-viewport overlays, empty layout, or missing overlay layout data.

Measured optimization gates matter: the 2026-05-19 FolderView optimization review did not ship a Task 7 production optimization because overlay and distant-state release gates lacked evidence. Future optimizations in those areas must produce same-machine before/after archives before being described as performance improvements.

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
  - `Common/PaneVisualState.h` (shared focused/inactive pane opacity policy)
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
- Programmatic path changes that replace or clear the NavigationView path MUST retire any active path-edit or full-path-popup native text host before hiding/resetting that child window. A refresh of the same path MAY preserve the edit host, but a different path or empty path must deactivate it so stale TextInput/UIA callbacks cannot navigate or paint against destroyed child HWND state.
- Embedded NavigationView instances MUST NOT use generic blur-reclaim behavior intended for the top-level pane address bar. Focus recovery after edit/popup dismissal remains owner-scoped so Find/results, file-operation popups, and other embedded hosts can close without stealing focus back to the pane.

## Current Item, Selection, and Location State Contract

### Core invariants

- A **displayed real item** is an item in the accepted model after provider enumeration, visibility rules, hidden-name exclusions, filtering, and sorting. Busy/error overlays, watermarks, host empty messages, and the standard empty-folder parent action are not real items.
- The **current item** is FolderView's persistent item cursor. It is independent from Win32 keyboard focus and from selection.
- Every nonempty displayed real-item model MUST have exactly one current item: `_focusedIndex` is valid, exactly one `FolderItem::focused` flag is true, and both representations identify the same item.
- An empty displayed real-item model has no current item, no selected items, and no range anchor. This is the only stable no-current state.
- Selection is an independent concrete set containing from zero through all displayed real items. The current item may be selected or unselected; moving, repairing, or restoring current MUST NOT implicitly select it.
- Every accepted model replacement MUST finish through one invariant-normalization operation that clears stale focused flags, establishes the resolved current item, and makes the current item visible under the normal viewport policy.

### Logical-location identity

The same-location versus navigation decision uses one provider-aware key:

1. logical file-system plugin ID;
2. stable provider-instance, mount, or connection context;
3. the logical folder key produced by the active provider's `FileSystemPathIdentity` contract.

Two folder presentations are the same logical location only when all three components are equivalent under the provider contract. Alternate path spellings, junctions, reparse paths, and remounted connections are not guessed by FolderView: they are the same only when the provider returns equivalent identity keys. If a provider cannot promise stable identity across recreation, FolderView scopes identity to the exact live instance context and exact ordinal path text; a recreated instance is navigation.

### Current-item resolution

For an accepted rebuild of the same logical location, including refresh, sort, filter, hide/show, or visibility-policy changes, resolve current in this order:

1. a validated one-shot host target, including successful create/rename intent or the specialized proven-removal result;
2. the surviving prior current identity;
3. the first surviving successor of the prior current under the resulting UI order;
4. the nearest surviving predecessor;
5. the first displayed item.

For an active keyed sort, successor/predecessor use the exact production directories-first comparator, sort direction, and stable tie-breaker against a non-actionable probe for the removed prior current. For Sort None, scan the prior accepted order forward and then backward for surviving provider-aware identities. The probe is resolver-only state: it is never selected, remembered, exposed through status/UI Automation, or accepted as a command target. Generic repair MUST NOT blindly reuse the old numeric index; if multiple neighboring rows disappear, it still chooses the first surviving successor before any predecessor.

A pure sort preserves current by identity. Passive focus memory MUST NOT override a surviving current during same-location changes. If filtering or hiding removes current, the resolved replacement becomes the remembered item; widening visibility later does not resurrect the older hidden current.

For navigation to another logical folder/provider/context, resolve current in this order:

1. a validated explicit navigation target;
2. the remembered current identity for the destination when it is displayed;
3. the first displayed item.

Navigation always starts with empty selection. Returning through Back/Forward restores current only, never selection.

### Bounded per-pane current memory

- Current memory is FolderView-owned, per pane, process-lifetime state and is not serialized.
- The key is the logical-location key above; the value is the remembered item component normalized and compared through the same provider identity contract.
- Each pane stores at most **512 entries** and at most **512 KiB** of owned UTF-16 provider/context/folder/item payload. Container overhead is additionally bounded by the entry cap.
- Insert/update and successful restore touch recency. Eviction is deterministic least-recently-used order with stable insertion order as the tie-breaker, and continues until both caps are satisfied.
- An individually oversized or unsupported identity is a no-cache event and MUST NOT evict the useful cache or fail navigation. `std::bad_alloc` remains fatal under repository policy.
- Switching roots/providers does not globally clear this cache. COM pointers, object addresses, HWND values, and other transient addresses MUST NOT be used as identity.

### Selection membership and persistence

- Only displayed real identities may be selected or used by selection-aware commands. Removed, filtered-out, hidden, placeholder, and stale identities are not latent selection.
- Same-location refresh/sort preserves selected surviving provider-aware identities. New identities start unselected. Removed/filtered/hidden identities leave selection immediately and an ordinary later reappearance remains unselected.
- A proven same-folder rename chain transfers selection only when its original identity was selected. The short-lived rename provenance is non-actionable and MUST NOT resurrect an ordinary disappearance/reappearance.
- Select All captures the displayed set at invocation time. If a later filter or visibility rule removes some selected items, those items leave selection; clearing the filter reveals them unselected, so Select All is no longer in effect.
- Select/Unselect masks, same-name/extension operations, inversion, and Hide Selected/Unselected mutate only the displayed identities present when the command runs. Showing names or widening visibility later does not apply the prior operation to returning identities.
- Successful Delete/Move removes exactly the proven-removed identities from selection. Failed, `S_FALSE`, cancelled, missing, and otherwise uncertain sources remain selected when their identity survives.
- Failed, cancelled, stale, or rejected enumeration results do not replace the accepted model and therefore preserve its current item and selection.
- Save/Restore Selection is the explicit user-requested exception to automatic no-resurrection. Save stores the selected displayed identities, or current-item fallback when selection is empty. Restore replaces live selection with saved identities that are displayed at that instant; identities excluded at restore time remain unselected until the user invokes Restore again after they become visible.

### Command, notification, rendering, and accessibility ownership

- Focus-only commands such as Enter/Open-current and F2/Rename use current. Selection-aware commands use selected displayed identities when selection is nonempty and otherwise fall back to current. Status selection summaries remain selection-only.
- Commands, menus, and drag gestures snapshot immutable provider-aware target identities before nested UI, asynchronous work, or drag startup. Later current/selection repair cannot retarget a captured operation.
- Current-item and selection notifications are independent and emit only when their corresponding observable state changes. Selection membership changes and selected-item metadata changes are distinguishable: membership changes cancel or replace obsolete size work, while a same-location metadata refresh for surviving selected folders preserves an explicit Space-size intent by recomputing the current surviving selection. Selection-only changes update selection statistics, status/command state, and size-work cancellation/recomputation but MUST NOT refresh current-item preview/detail consumers or change current ownership.
- Preview content follows current-item/path changes only. Selection-only input MUST NOT reload an unchanged current-item preview.
- Current ownership epochs change only when explicit user input changes current identity. Selection-only changes, including empty-background selection clearing, do not change current ownership and do not invalidate an otherwise matching pending Delete/Move rehome. Provider, folder, and sort epochs and exact-per-source-`S_OK` proof remain required.
- Rendering invalidates the old and new current bounds and cannot retain a stale focused flag. The active/inactive current border remains independent from selection fill.
- UI Automation selection exposure is selection-based. Current is not reported as selected merely because it is current. The separate empty-folder parent action, if exposed, uses Invoke rather than Selection/SelectionItem semantics.

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
- Horizontal thumb tracking may follow raw pixel/DIP offsets while the user is
  dragging, but thumb release MUST snap to the nearest valid column scroll stop.
- Hit testing treats the leading gutter before column 0 as part of column 0 so
  the first visible strip remains clickable. Gaps between later columns remain
  empty hit-test space.
- Visible body item snapshots used by rendering, hit testing, keyboard
  navigation, and UI Automation queries MUST be built as per-call immutable
  values or otherwise protected from concurrent mutation. Accessibility
  navigation can arrive re-entrantly or from UIA-owned threads, so these paths
  must not return references/views into a shared mutable visible-item cache.

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
- Filename hover tooltips are not implemented and are not part of the current FolderView contract; the optional product decision is recorded in `Specs/Plans/WIP/Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md`.

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
- For local files, the pane requests shell thumbnails asynchronously for visible items only. The normal visible thumbnail path MUST use cached-only shell lookup (`SIIGBF_INCACHEONLY`) and MUST NOT issue provider-allowed shell thumbnail lookups. If cached shell thumbnail extraction fails for a likely WIC-supported image extension, the worker decodes a bounded first-frame thumbnail through WIC before falling back to the normal icon.
- Product decision 2026-07-04: ship the cached-only visible path. Cold-cache non-WIC formats may initially show the normal icon where older builds produced real previews. Background provider-allowed enrichment is a required follow-up tracked by `Specs/Plans/WIP/FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md`; it must run after the fast visible path and must not weaken the cached-only first-paint contract.
- Shell and WIC thumbnail extraction are limited to local-shell-backed file systems. Virtual, remote-plugin, archive, cloud-plugin, or otherwise non-local file systems use the normal fallback icon path and must not perform native shell/WIC thumbnail extraction for provider paths.
- WIC fallback decoding MUST reuse one WIC imaging factory per enumeration worker and release it before that worker uninitializes COM. Thumbnail workers that use WIC are expected to have COM initialized as MTA.
- WIC fallback decoding MUST reject oversized source frames before decode; the current cap is `64 MP` source pixels. This is a safety bound, not a quality target.
- Any provider-allowed shell thumbnail work used for diagnostics or approved background enrichment MUST be abandonable and bounded: it captures only values (`HWND`, generation, item index, path, target size), never captures `this`, initializes COM on the worker, waits only for a short deadline, never blocks navigation or close on a join, and posts late results through the normal payload path where stale `HWND`/generation/batch checks drop obsolete work.
- Thumbnail bitmap UI-thread payloads MUST account for pending work exactly once. Normal thumbnail worker posts increment `pendingBitmapCreates` before posting and the UI-thread apply path decrements only for payloads that own that increment. Stale batch or stale enumeration-generation payloads MUST return before decrementing the current batch's pending count, and late provider/background-enrichment payloads posted after the deadline MUST be marked uncounted so they can apply or drop without consuming another request's pending slot.
- Thumbnail diagnostics MUST expose and/or emit cached shell hit/miss and provider-allowed activity: `thumbnails.cached_extract_us`, `thumbnails.shell_cache_hit_count`, `thumbnails.shell_cache_miss_count`, `thumbnails.shell_provider_allowed_count`, `thumbnails.shell_provider_timeout_count`, and `thumbnails.close_to_idle_us`. Shell thumbnail stat increments that update a debug snapshot counter and emit a matching perf counter MUST route through the shared thumbnail stat helper so the two telemetry surfaces cannot drift.
- Until a thumbnail is ready, or when neither shell nor WIC can provide one, the normal folder/file icon is rendered in the same larger visual slot.
- Folder/file icon bitmap caching MUST keep separate entries for the shell image-list size class selected for the current target DIP size. Returning from thumbnail mode to Brief/Detailed/Extra Detailed MUST NOT reuse thumbnail-mode jumbo icon bitmaps in the normal `16 DIP` icon slot.
- Thumbnail-size changes while thumbnail mode is active MUST clear size-dependent fallback icon bitmaps as well as thumbnail bitmaps, reset dependent layout state, and requeue normal icon loading before thumbnail work so fallback slots do not stretch stale icons from the previous target size.
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

Folder View and its per-pane status bar use the shared `CompactDetails` normalized-metadata display
profile. A positive valid Windows FILETIME is converted with the current Windows local time zone and
rendered as fixed `YYYY-MM-DD HH:MM`; zero, negative, or invalid values yield no time text. Compact
attributes use the exact `RHSACETOP` order (read-only, hidden, system, archive, compressed, encrypted,
temporary, offline, reparse point), or `-` when none apply. Directory and sparse-file bits are not
shown by this compact profile; compressed remains visible. Sizes continue to use `FormatBytesCompact`,
including `0 B` for a zero-byte file.

Context owns omission and localization. Folder item details retain their fixed separator slots and
omit the size slot for directories. The status bar omits an unavailable timestamp by selecting the
matching localized resource form; an unresolved directory size uses the localized unknown-size text
instead of inventing a numeric value.

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

**Pointer target semantics:**
- The current item, multi-selection set, and transient pointer/drag target are independent state dimensions. Clearing a pointer target or selection MUST NOT remove the current item from a nonempty display.
- Every left-button-down that hits a displayed real item MUST make that item current during button-down handling, before drag initiation or later command routing observes the gesture. For Shift/Ctrl+Shift with no valid anchor, FolderView MUST first capture the pre-press current item as the inclusive range anchor, then make the hit item current; both anchor and clicked endpoint belong to the range.
- A plain left-button-down on an unselected item MUST replace the prior selection with that item before drag detection or command routing reads the selection. A plain left-button-down on an already-selected item MUST preserve the existing multi-selection while moving current to the pressed item.
- Ctrl+Click MUST toggle only the clicked item. Shift+Click MUST replace selection with the inclusive anchor-to-clicked range. Ctrl+Shift+Click MUST add that inclusive range to selection. A valid prior anchor wins; otherwise Shift/Ctrl+Shift uses and includes the valid current item captured immediately before the press.
- A pointer right-click on an unselected item MUST establish that item as the sole target before the context menu is built. A pointer right-click on an already-selected item MAY preserve the existing multi-selection.
- A left-button miss on empty background, whether unmodified or combined with Ctrl, Shift, or Ctrl+Shift, MUST clear selection, clear the transient pointer/drag target, disarm any armed drag source, retain current, and reset the range anchor to current. The modifiers have no range/toggle endpoint on a miss. Current ownership does not change.
- A pointer right-click miss MUST perform the same selection-clear/retain-current/no-ownership-change transition, then route to the background context menu with no item pointer target. Item-only menu commands MUST NOT infer a target from the retained current item.
- A keyboard context-menu request is not a pointer miss. When a selected/current item exists, it uses the selected-or-current command fallback and MUST NOT hit-test a synthetic client-center point or clear selection. When the pane contains no current item, it loads the background menu; item-only commands are absent.
- Empty-background movement MUST NOT start a drag from the retained current item or a stale selection: a later drag requires a fresh button press that hits an item and establishes a new pointer/drag target.
- An item left-button-down MAY arm a potential drag from an immutable post-press source snapshot: the selected displayed identities when selection is nonempty, otherwise the current item. The pressed item must belong to that resolved source; if a modifier gesture leaves it outside the source, no drag is armed. Drag starts only after pointer movement leaves the effective system drag-threshold rectangle centered on the press point. The item's visual bounds MUST NOT add a second threshold. Release inside the threshold rectangle completes the click without dragging. Release, Escape, cancellation, capture loss, teardown, navigation/model replacement, or any new press MUST disarm the prior gesture. Escape also clears selection, retains current, and resets the range anchor to current; Quick Search consumes the same Escape for search exit but still disarms the armed pointer gesture.
- A double-click MUST apply the normal first-press current/selection rules and activate the resulting current item; it MUST NOT reuse an armed source from an earlier gesture.
- Destructive context-menu commands MUST execute only when the selected/current provider-aware target set still matches the immutable snapshot captured before the nested menu loop. Generic disappearance repairs current through the same-location successor/predecessor contract above and MUST NOT retarget the captured command. Host-owned Delete/Move remains the more specific exception: it captures its source intent and ownership epochs, accepts only exact per-source `S_OK` removal proof, and resolves replacement only against a newer accepted enumeration after the old current identity disappears. The replacement remains unselected unless it was already a surviving selected identity.

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
    - The localized **Go to parent** row is a separate empty-state Invoke action, not a `FolderItem`, current item, selection, range anchor, remembered item, drag source, or item-command fallback.
    - Enter while the standard empty-folder action is active, double-click anywhere in that surface, and Backspace invoke the existing `NavigateUp()` route. Root/connection-root callback behavior remains unchanged.
    - The action's pane-focus visual spans the current pane row width and uses the current display mode's row height, but it is never exposed as selected or as FolderView current. Its metrics MUST be recomputed from the current client width, DPI, icon size, and display-mode text-line heights and MUST NOT inherit prior item geometry.
    - Filter-empty, explicit host-message, busy/cancel, error, and background-watermark surfaces do not acquire this action.

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
- Every filter/visibility/hide operation removes excluded identities from live selection immediately. Clearing the filter, showing hidden/system items, or showing hidden names later reveals those identities unselected; there is no latent selection.
- Select All is a snapshot of the displayed identities at invocation time, not a sticky mode. Select All followed by filter/hide and then unfilter/show preserves selection only for identities that stayed continuously displayed; returning identities are unselected.
- While a pane filter is active **or** the hidden-names set is non-empty, FolderView SHOULD render a very subtle watermark glyph in the background (Sego UI Symbol `0xE71C`) so the user can tell the pane is filtered/hidden.
- When the pane filter is active and no rows are visible, the filter state has priority over the generic Empty folder placeholder:
  - FolderView MUST suppress the generic Empty folder centered UI and its parent-row placeholder so the pane does not imply that the directory itself is empty.
  - FolderView MUST render the full-pane funnel watermark/placeholder for the filtered-empty result, unless an error/busy overlay is visible.
  - For explicit host empty-state messages that remain visible while filtering, FolderView MUST avoid overlapping watermarks and SHOULD use a small filter badge instead.

**Enumeration Contract (Plugin Only):**
- The host obtains an `IFileSystem` instance via the plugin factory (`RedSalamanderCreate(..., pluginId, ...)`) and uses it as the only source of directory entries.
- Each enumeration calls `IFileSystem::ReadDirectoryInfo(path, info.put())` to obtain an `IFilesInformation` result object.
- The returned `FileInfo` buffer is traversed via `NextEntryOffset` (preferred) to build `FolderItem` entries.
- `IFilesInformation::GetCount()` is an untrusted allocation hint. FolderView MUST clamp any reserve derived from it by the validated buffer size and a fixed implementation ceiling, and allocation/reserve work MUST remain inside the enumeration worker's exception boundary.

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

**File Operations artifact presentation:**
- FolderView never hides, filters, or removes an item merely because it is a Possible File
  Operations artifact. The item stays in its ordinary sort position and remains subject only to the
  user's normal hidden/system, wildcard, and hide-name choices.
- The shared artifact classifier supplies an orthogonal row state. Possible renders a warning overlay
  and `Possible interrupted-operation artifact`. This semantic remains available to accessibility
  even when the current display mode has no details line.
- A badge or pane cache is not authority. Possible is a recognized generated-name shape only and
  never proves ownership, cleanup, or recovery authority. Folder refresh or watch change re-queries
  classification.
- Each enumeration loads one empty artifact-classifier snapshot. Because the production journal and
  durable claims are retired, no endpoint/parent claim index is populated or consulted. Only a
  generated artifact name shape enters the no-follow projection probe; ordinary names perform no
  artifact capability, identity, or bind work. A future durable claim source must define its own
  authoritative schema before keyed lookup can return.
- When an artifact row is focused with no selection, the pane status bar appends the same localized
  Possible classification text used by Find. The artifact context submenu exposes read-only
  `Inspect` and `Reveal in pane`; Inspect forces the built-in host viewer so a configured external
  viewer cannot turn the inspection action into an unguarded launch. `Reveal in pane` is the read-only
  Open-location presentation. There is no Resume, Roll back, cleanup, or recovery action.
- Selection, Properties, Reveal/Open location, and the host read-only viewer are inspection. Copy,
  external Open/Edit, Move, Rename, Delete/Recycle, overwrite, write-capable commands, drag, and
  clipboard operations touching a Possible row route through the shared exact-set artifact warning
  before admission. A grant applies only to the captured identities and is revalidated before use.
- Possible rows remain selectable. A normal explicit file command may proceed only through the
  Possible-artifact warning; consent never becomes automatic cleanup authority.

**Icon Rendering:**
- Uses IconCache system image lists (`SHIL_SMALL`/`SHIL_LARGE`/`SHIL_EXTRALARGE`/`SHIL_JUMBO`) and selects the **optimal** list size based on the target icon DIP size and current DPI (FolderView default is 16 DIP list-mode icons).
- IconCache MUST choose the smallest shell image-list source that is at least as large as the target physical pixel size; do not upscale 16px shell icons on high-DPI displays.
- FolderView icon bitmap draws MUST use nearest-neighbor interpolation only when the source bitmap and destination physical pixel size match exactly; scaled shell icons use linear filtering to avoid chunky edges.
- IconCache D2D bitmap normalization MUST premultiply translucent BGRA pixels and, for shell icons that carry RGB data but no alpha channel, apply Windows icon AND mask semantics: black mask pixels are opaque and non-black mask pixels are transparent.
- Fallback chain: **optimal → remaining sizes** (best-effort quality preservation)
- Icons cached in IconCache component (LRU cache, 2000 icons ≈18MB)
- Async loading with viewport prioritization: visible items first, offscreen queued
- Per-file icon extraction for .exe, .dll, .ico, .lnk, .url, .cpl, .scr, .msc, and .ocx (embedded icons)
- Extension-based caching for common file types (bypasses Shell API on cache hit)
- Per-file icon-index lookups for items flagged `FILE_ATTRIBUTE_OFFLINE`, `FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS`, or `FILE_ATTRIBUTE_RECALL_ON_OPEN` MUST use shell attribute-only lookup (`SHGFI_USEFILEATTRIBUTES`) instead of live path lookup so dehydrated cloud/offline placeholders are not recalled. Each avoided live lookup emits `icons.recall_avoided_count`.
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
- `CFSTR_PREFERREDDROPEFFECT`: Suggests copy vs. move. FolderView paste MUST honor `DROPEFFECT_MOVE` as a real file-operation move and treat `DROPEFFECT_COPY`, missing metadata, malformed payloads shorter than `sizeof(DWORD)`, or unsupported metadata as copy.
- Private FolderView drop data and `CF_HDROP` are untrusted. Readers MUST validate `GlobalSize`, header/offset/alignment, wide encoding, required terminators, count-derived minimum bytes, per-path length, aggregate path length, and a fixed maximum item count before reserving or constructing strings. Clipboard readers MUST apply the same item/path/aggregate bounds; CF_HDROP writers MUST reject arithmetic overflow before `GlobalAlloc`.
- A paste carrying `DROPEFFECT_MOVE` captures the complete item list and clipboard sequence. The shared File Operations engine publishes the accepted task and creates a gated worker before asking FolderView to clear that exact sequence; only after the clear attempt does it release the worker. Thread/queue admission failure leaves the clipboard untouched. Clear failure warns but does not discard the accepted task, and the engine's bounded sequence ledger rejects another paste of that exact accepted sequence. A successfully cleared cut payload is never restored, including when the result is `Copied; source kept`, partial, canceled, failed, or indeterminate. Copy clipboard content remains repeatable.
- Completed retained-source actions select exact display names after the next matching enumeration. `Cut retained items again` is an explicit new clipboard write and is available only for a complete known retained set; immediately before selection/cut, FolderView's owner rebinds every local source no-follow and requires its captured object identity. One missing or replaced item aborts the entire action with a warning.
- Completion focus and cache updates consume per-item result axes. FolderView removes or advances
  past a source row only for exact `SourceDisposition::Removed`; Retained and Unknown rows remain and
  refresh. Destination folders refresh only for Published or Unknown publication, so a known
  NotPublished conflict does not masquerade as a completed transfer.
- When a `FileOperationRequestCallback` is installed by `FolderWindow`, FolderView copy/move/paste/drop paths MUST delegate copy and move work without showing a local pre-confirmation. The shared File Operations layer owns the single OK/Cancel confirmation for those delegated operations. Clipboard paste and folder-picker move MUST use the same delegated route and MUST NOT pre-grant overwrite, replace-read-only, or continue-on-error flags.
- Clipboard paste, internal drag/drop, and external OLE drop capture one immutable File Operations options snapshot, including verification and link policy, and submit qualified source/destination endpoints. No worker rereads live Preferences or provider JSON.
- Paste Shortcut is local-file-system only. FolderView reads the clipboard on the UI thread, then creates `.lnk` files on value-only background work that captures copied source paths, target folder, path-visit generation, original filesystem identity, request id, and `HWND`, but never captures or dereferences `this`. A FolderView MUST serialize Paste Shortcut requests: if one worker is in flight, later invocations queue the already-captured clipboard payload and run after completion. The worker initializes COM as MTA before calling shell-link APIs and posts completion through `PostMessagePayload`.
- Paste Shortcut shortcut creation MUST be collision-safe. Each deterministic `GenerateShortcutPath` candidate is probed and saved without replace semantics; if a racing creator wins the same slot (`ERROR_FILE_EXISTS` / `ERROR_ALREADY_EXISTS`), the worker retries the next candidate. Short and long final paths share the same temp-file-plus-move save path. The shell-persist temp path remains reserved between `GetTempFileNameW` and `IPersistFile::Save`; the placeholder is not deleted and exposed to name stealing.
- Paste Shortcut completion MUST run on the UI thread. Completion always emits `clipboard.paste_shortcut_worker_us`; cache invalidation for `result.targetFolder` uses the captured original filesystem even if the pane changed provider, while current-view refresh and focus restoration additionally require that provider identity, target path, and path-visit generation still match. A failed first completion post is retried from a preserved result payload; request-id matching, a stale timeout, and `WM_NCDESTROY` reset prevent a lost or late completion from wedging or clearing a newer request. Failure reporting via the Paste Shortcut pane alert remains unconditional when the UI receives completion. A stale completion MUST NOT navigate back to the old folder or focus an item there.
- Missing `FileOperationRequestCallback` is a host wiring failure for normal UI paths. FolderView MUST fail visibly and log the error instead of silently running a direct plugin copy/move/delete with different conflict/progress semantics. Direct plugin fallback is reserved for explicit no-host/test scenarios with an opt-in test hook.

**Drag Initiation:**
```cpp
// After an item press moves outside the system drag-threshold rectangle:
const auto sourceItems = GetSelectedItemsOrCurrentItem();
wil::com_ptr<IDataObject> dataObj = CreateDataObject(sourceItems);
wil::com_ptr<IDropSource> dropSource = CreateDropSource();
DWORD effect;
DoDragDrop(dataObj.get(), dropSource.get(), 
           DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK, &effect);
```

- Drag-source resolution uses selected displayed items when selection is nonempty and otherwise falls back to the current item, matching selection-aware command targeting.
- The press must hit an item contained in the resolved post-press source snapshot. Empty-background movement and a modifier press that leaves the hit item outside the source cannot start a drag.
- The source identities are snapshotted for the gesture before OLE drag startup. Later selection/current repair cannot retarget an in-progress drag.

**Visual Feedback:**
- System-provided drag image (ghost icon + file count badge)
- Cursor changes based on drop target (copy/move/no-drop)

#### Drop Target (Dragging INTO FolderView)

**COM Interface:**
- `FolderView::DropTarget`: Implements `IDropTarget`

**Drop Operations:**
- **Copy**: Ctrl key held during drop
- **Move**: Shift key held, or the no-modifier default when every local source and the destination are on the same volume and MOVE is allowed
- **Copy across volume/provider boundaries**: Default when no modifier is held and same-volume MOVE cannot be established
- **Link**: Ctrl+Shift keys held
- **Cancel**: Escape key
- Delegated copy/move drops follow the same confirmation ownership as clipboard paste: exactly one shared File Operations prompt, not an additional FolderView prompt.
- An external OLE drop that queues an asynchronous MOVE MUST report COPY to the external source at drop return. Queuing is not proof that MOVE committed and MUST NOT authorize early source deletion. Internal FolderView gestures retain their provider-qualified source identity for the host operation.

**Drop Validation:**
- The screen drop point MUST be converted to client coordinates. A hovered directory is the destination; pane background uses the currently displayed folder.
- Drop, paste, and cross-pane copy/move destinations MUST be fully enumerated. A destination still loading is rejected with localized pane feedback rather than silently re-targeted.
- Every source MUST be rejected when the resolved destination is its current parent, the same path, or its own descendant.
- External `CF_HDROP` sources are local filesystem paths. When the destination pane is virtual, the host MUST use the builtin local provider as the source and the pane provider as the destination, subject to the cross-filesystem capability gate; it MUST NOT ask the destination provider to read local paths.
- Validate the resolved source/destination providers support the operation.
- Highlight drop target folder during hover
- Show appropriate cursor (copy/move/no-drop arrow)
### 3. Context Menus

**Trigger:** Right-click on file/folder or selected items

**Localization requirement:** the context menu template is defined in `.rc` resources (see `Specs/Core/Core_Localization.md`) and loaded at runtime.

**Resource menus:**

- `IDR_FOLDERVIEW_ITEM_CONTEXT` in `RedSalamander/RedSalamander.rc`
- `IDR_FOLDERVIEW_BACKGROUND_CONTEXT` in `RedSalamander/RedSalamander.rc`

**Item target:**

- Open; Open With…; Calculate Occupied Space
- separator
- Cut; Copy; Paste
- separator
- Move…; Delete…; Rename…
- separator
- Operation artifact submenu when the focused row is a Possible artifact; separator
- Properties

**Background target:**

- Paste
- New > Folder…, Edit New File…, then dynamic Shell New templates
- separator
- Refresh; Calculate Occupied Space

**Runtime behavior:**
- A pointer invocation on an item targets the preserved selection when the hit item is already selected; otherwise it establishes the hit item as the sole target. A keyboard invocation targets the selected set, or the current item when selection is empty; with no current item it loads the background resource.
- A pointer miss loads only the background resource. Retained current identity is not an item target and item-only commands are absent rather than disabled.
- `Cut` is enabled only when the resolved item target can publish local built-in-file-system `CF_HDROP` paths. Providers that the clipboard implementation rejects (including Dummy, S3, MTP, and cloud providers) show Cut disabled rather than a working-looking command that warns after invocation.
- `Paste` preflight uses the same non-blocking clipboard-format check as the paste path and also requires a live, enumerated destination folder. The file-operation admission path remains authoritative for provider capabilities and revalidates at execution; the popup MUST NOT perform blocking provider I/O merely to predict that later result.
- The optional Operation artifact submenu is inserted immediately before the separator and final `Properties` action. It is never appended after Properties or debug-only suffixes, and a failed submenu/separator insertion rolls back partial composition.
- Selection-aware destructive commands compare the post-menu target identities as a multiset, preserving duplicate cardinality; a display-name membership check that can accept a duplicate or substituted target is insufficient.
- Menu rendering uses the active `MenuTheme` (owner-draw for themed background/selection colors).
- DxUI context menus opened from right-click or keyboard context-menu commands MUST share the common `DxUi::ContextMenu::Show(...)` input contract: pointer hover highlights visible rows, enabled item clicks invoke immediately, outside clicks and Escape dismiss, arrow keys move the keyboard highlight, Enter/Space invoke, and focus returns to the owning pane when the menu exits.
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
- **Space**: In normal item-navigation mode, capture the `Current item`, toggle its selection, request selection-size recomputation from the complete post-toggle selected set, and then move `Current item` to the next displayed item when one exists. An initially unselected current item becomes selected and contributes its known file size or asynchronous folder-subtree size; an already-selected current item becomes unselected, pending work for its contribution is canceled, and any stale completion is excluded from the selected total. A reusable cached folder size may remain cached. The move does not wrap, and the destination item is not implicitly selected.
- **Insert**: Toggle selection of the captured `Current item` and then move to the next displayed item without requesting explicit folder-subtree size computation. The move does not wrap, and the destination item is not implicitly selected.
- **Alt+Up / Alt+Down**: Move current to the previous/next selected displayed identity, wrapping within the selected set. Preserve selection, set the destination as range anchor, and change current ownership only when the current identity actually changes. *(Default chord binding is settings-backed.)*
- **Ctrl+Shift+<key left of Backspace>**: Select all items with the same extension as the focused item (adds to selection).
- **Ctrl+Shift+<key right of 0>**: Unselect all items with the same extension as the focused item (removes from selection).
- **Ctrl+A**: Select all items *(default chord binding is settings-backed)*
- **Ctrl+Click**: Toggle individual item selection
- **Shift+Click**: Range selection from anchor to clicked item
- **Shift+Arrow/Home/End/Page**: Replace selection with the inclusive anchor-to-endpoint range and make the endpoint current. Preserve a valid anchor; otherwise capture and include the pre-move current item as anchor.
- **Ctrl+Shift+Arrow/Home/End/Page**: Add the same inclusive range to selection and make the endpoint current. Preserve a valid anchor; otherwise capture and include the pre-move current item as anchor.
- **Ctrl+Shift+Click**: Add the inclusive anchor-to-clicked range and make the clicked endpoint current during button-down handling.
- **Shift+Home/End**: Extend selection to start/end
- **Esc**: Clear selection


**Action Keys:**
- **Enter**: Open focused item (folder navigates; file invokes host open hook which may mount a virtual file system or fall back to `ShellExecute`)
- **Delete**: Delete selected items (with confirmation)
- **F2**: Rename focused item

F2 submits a one-step `RenamePlan` with `RenameOrigin::InlineRename` to File Operations workers; it never
calls provider `RenameItem` on the UI thread. The inline editor is the only initial confirmation.
Before publication, central admission validates the proposed child through the retained typed provider
route and stores only the provider-owned joined path and collision key. Provider Invalid preserves its
exact HRESULT; missing, unsupported, or malformed name methods fail closed. The worker re-queries the
same composite contract after interlock waits and immediately before conditional mutation, so route
drift cannot retarget the rename. FolderView never substitutes Windows name rules or host path joining.
An invalid/unsupported request rejected before task publication, or missing host callback wiring,
shows the localized pane operation-error overlay. Once a task is published, its card is the sole
failure presentation and FolderView does not add a duplicate pane alert.
Inline F2 bypasses global Queue/Parallel admission but honors overlapping-root interlocks. Conflict,
waiting, failure, partial, canceled, or indeterminate state reveals the task immediately. A clean
running task is revealed after 500 ms; a clean success before that deadline creates no task card or
completed row, still refreshes the pane, and retains focus by renamed identity. Once revealed, the
card follows normal task rules and never becomes silent. Tests use a fake clock or zero delay.
Batch Rename and any F2 action that opens Batch Rename use ordinary queue/card behavior.
- **Backspace**: Navigate to parent folder
- **F3/F5/F6/F7/F8**: FolderWindow-global operations (view/copy/move/mkdir/delete) per `Specs/UI/UI_CommandMenuKeyboard.md`.

**Clipboard Keys:**
- **Ctrl+C**: Copy selected items to clipboard (or `Current item` when selection is empty) *(default chord binding is settings-backed)*
- **Ctrl+X**: Cut selected items to the shell clipboard (or `Current item` when selection is empty) *(default chord binding is settings-backed)*
- **Ctrl+V**: Paste from clipboard to current folder, moving instead of copying when the shell clipboard carries `Preferred DropEffect = DROPEFFECT_MOVE` *(default chord binding is settings-backed)*
- Ctrl+X followed by Ctrl+V in a normal hosted pane MUST request exactly one shared move confirmation before the File Operations task starts.

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
- Same-folder refreshes MUST preserve active incremental search state and query text. FolderView clears incremental search on actual path replacement/clear, not on a refresh where the displayed folder and current folder still identify the same location.
- `WM_KILLFOCUS` MUST exit incremental search only when focus moves outside the FolderView subtree. Null focus transitions and focus repairs that leave focus on the FolderView or one of its children MUST preserve the active search session.

### 5. Multi-Selection

**Selection Modes:**
1. **Plain left-button-down**: Make the hit item current immediately. If it is unselected, replace selection with that item; if already selected, preserve the existing selection.
2. **Ctrl+Click**: Toggle clicked item selection, keep others
3. **Shift+Click**: Replace selection with the inclusive range from anchor to clicked item. With no valid anchor, capture the pre-press current item as anchor and include it before moving current to the clicked endpoint.
4. **Ctrl+Shift+Click**: Add the same inclusive range to the existing selection; use the pre-press current item when no valid anchor exists.

**Visual Feedback:**
- On folder entry (after enumeration), FolderView sets the `Current item` to the first item (or nearest preserved focus), and starts with **no selection**.
- On refresh of the currently displayed logical location, including refreshes triggered by `DirectoryInfoCache` callbacks, preserve selection for every surviving provider-aware identity even when metadata/layout/icon state changes. Remove identities that are no longer displayed. A selected identity may transfer only through a proven chained same-folder rename hint to its final identity; an unselected original MUST NOT become selected merely because it was renamed. Without proven rename provenance, disappearance is removal and an ordinary later reappearance is unselected.
- A host-owned delete or move treats only an exact per-source `S_OK` result as proof that source was removed; aggregate success, `S_FALSE`, cancellation, missing/not-found, or failure MUST NOT prove removal. Source identity comparisons follow the active provider's path case-sensitivity contract. Once a newer accepted enumeration no longer contains the old focused name, the successful proof feeds the same canonical resulting-UI-order successor/predecessor resolver defined above. Every source proven removed is excluded as a repair candidate even if a stale accepted provider snapshot still contains its identity. The specialized result MUST NOT derive a replacement by applying a captured display index to the pre-sort enumeration payload. If the accepted nonempty model contains no candidate other than identities proven removed, the specialized intent nominates no target and the normal nonempty-model invariant fallback applies until a fresher enumeration arrives. Folder and file removal share this rule; the replacement remains unselected unless it was already part of a surviving selection.
- Host-owned Move uses the same removal-focus path as Delete. `StartFileOperationFromFolderView` tracks `FILESYSTEM_MOVE` as well as `FILESYSTEM_DELETE`. `CommandMoveToOtherPane` tracks the **source** pane even when the destination is the other pane. Fail-closed epoch and status checks stay the same: a user current-identity change, navigation/provider change, sort change, or a non-`S_OK` current-source outcome MUST NOT apply the specialized rehome.
- `BeginRemovalFocusTracking` captures `folderPathGeneration`, the focused display name, and only the small requested source-name list plus source indices needed to correlate completion outcomes. It MUST NOT copy every visible `displayName`, capture per-visible-row ordering, or build a packed ordered-name buffer on the UI thread. After proof, current repair reuses the existing probe and production sort comparator rather than maintaining a second removal-specific ordering walk. Begin tracking MUST emit `folder.removal_focus.begin_us` with `value0` = visible item count and `value1` = requested source count. Focused x64 Release evidence on machine `4cb089111a23` measured 490 µs for 4096 visible items and one requested source; the validated archive is `Specs/TestRuns/4cb089111a23/Commands/2026-08-19_183856`. This is a diagnostic sample, not a hard performance budget.
- Pending removal-focus intent is valid only while provider identity, folder path, sort epoch, and current-ownership epoch still match. A stale enumeration that still contains the old current item retains the intent for a later generation. A user gesture that changes current identity, navigation/provider change, or sort change invalidates it; selection-only changes, including empty-background clearing that retains current, do not. When refreshes coalesce, the newest applicable committed sequence wins and older intents cannot steal current.
- Same-folder refreshes MUST emit one aggregate metric row per refresh for `folder.refresh.preserve_count`, `folder.refresh.rebuild_count`, `folder.refresh.selection_preserve_count`, `folder.refresh.rename_transfer_count`, `folder.refresh.debounce_delay_ms`, and `folder.refresh.enumeration_count`; they MUST NOT emit per-item refresh rows. `folder.refresh.request_to_paint_us` is emitted once after the next successful present following the refresh result for the same current enumeration generation and uses a dedicated pending refresh slot, not the shared input/navigation-to-paint slot.
- A ready pending refresh-to-paint metric MUST be discarded, not retained for a later unrelated frame, when the paint path cannot present: missing render target/swap chain/target bitmap fallback, no-target render exit, `EndDraw` failure, `Present1` failure, or legacy `Present` failure. Starting a newer refresh resets any existing pending refresh-to-paint slot before recording the new one.
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
- When a top-level `FolderWindow` moves between monitors with different DPI/scale, it MUST notify both child `FolderView` instances of the new DPI before resizing/repositioning child windows. This keeps row metrics, Direct2D context DPI, icon target sizes, and child `WM_SIZE` work on the same scale.
- `FolderView` MUST treat DPI changes, client-size changes, successful swap-chain buffer resize, and new Direct2D target creation as full-client repaint boundaries. The next paint must expand the dirty rectangle to the full client and keep doing so until a successful present clears the pending full-render flag.
- The full-client repaint boundary prevents newly allocated or DPI-retargeted back buffers from exposing black/uninitialized regions or stale scaled content while dragging the window between displays such as 150% and 100% scale.

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

**Teardown note:** windows that receive payload messages must call `InitPostedPayloadWindow(hwnd)` on create and
`DrainPostedPayloadsForWindow(hwnd)` in `WM_NCDESTROY`. Posted nonzero `lParam` values are opaque registry tokens,
not payload pointers. Draining fences new posts, invalidates the window's tokens, and deletes registered storage
exactly once; a stale token left in the thread queue or observed after numeric `HWND` reuse is rejected by
`TakeMessagePayload<T>` and must be ignored.

### IconCache Integration

**Architecture:**
- **Singleton instance** shared across all FolderView windows
- **LRU cache** of D2D bitmaps (2000 icons; worst-case ≈18MB at 48×48 BGRA, smaller at 16×16 or 32×32 depending on target DIP size/DPI)
- **Extension-to-iconIndex mapping** for instant lookups (bypasses Shell API)
- **Per-file whitelist** for unique icons (.exe, .dll, .ico, .lnk, .url, .cpl, .scr, .msc, .ocx)
- **3-level fallback** chain: selected optimal size → other sizes (from `SHIL_SMALL`/`SHIL_LARGE`/`SHIL_EXTRALARGE`)
- Extension-association misses fall back to the generic shell file/folder icon index instead of leaving `iconIndex = -1`.
- Live per-file shell icon lookup failures (`useFileAttributes=false`) are treated as
  transient and use a per-path bounded negative-cache backoff: `250 ms` after the
  first failure, doubling after consecutive failed retries to a `4 s` cap. A retry
  after the active window re-enters `SHGetFileInfoW`; any success replaces the
  negative entry with the positive icon index and resets the failure history.
- Attribute-mode computed icon lookup failures (`SHGFI_USEFILEATTRIBUTES`) may use the
  bounded negative cache because those failures are deterministic for the requested
  attributes.
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
    item.iconIndex = IconCache::GetInstance().QuerySysIconIndexForPath(fullPath.c_str(), item.fileAttributes, false).value_or(-1);
}

// 3. Convert icon index to D2D bitmap on UI thread (cached per D2D device)
auto bitmap = IconCache::GetInstance().GetIconBitmap(item.iconIndex, _d2dContext);
```

**Viewport-Aware Loading:**
- Visible items processed first (high priority)
- Offscreen items queued (low priority)
- Icon bitmap extraction runs through the background icon-loading pipeline; do not add additional extraction concurrency unless fresh same-machine evidence shows a material queue/extract bottleneck.
- Telemetry logged: total requests, visible requests, cache hits, extracted count, duration. IconCache lock diagnostics MUST be thresholded slow-path rows (`iconcache.lock_wait_slow_us` / `iconcache.lock_hold_slow_us`) with the IconCache stage in `detail`; raw per-lock wait/hold rows are too noisy for archived perf runs and MUST NOT be emitted. Shell icon lookup timings (`iconcache.shgetfileinfo_us`) MUST identify the lookup kind in `detail` (`association`, `path_attributes`, or `path_live`) and carry file attributes plus SHGFI flags in the value fields. Every actual failed live-path shell lookup emits `iconcache.path_live_lookup_failed_uncached`; a live negative-backoff hit emits both `iconcache.path_failed_lookup_cache_hit` and `iconcache.path_live_lookup_failure_cache_hit`; stored or refreshed negative entries emit `iconcache.path_failed_lookup_cached`. A slow failed shell query MUST NOT downgrade a successful concurrent icon-index result; losing duplicate store races emit `iconcache.duplicate_path_query_race`.
- The FolderView icon-pipeline metric contract includes `FolderView.ExecuteEnumeration.IconIndex.QueryExtensions`, `FolderView.ExecuteEnumeration.IconIndex.QueryPerFileIcons`, `FolderView.ExecuteEnumeration.IconIndex.BuildPerFilePaths`, `FolderView.IconLoading.ProcessQueue`, `FolderView.IconLoading.BatchUpdate`, `FolderView.IconLoading.BitmapConversion`, `icons.queue_wait_to_dequeue_us`, `icons.extract_us`, `icons.batch_update_scan_us`, `iconcache.shgetfileinfo_us`, `iconcache.lock_wait_slow_us`, `iconcache.lock_hold_slow_us`, and `icons.recall_avoided_count`.

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
- Recoverable Direct2D/DXGI device loss from `ID2D1DeviceContext::EndDraw`, `IDXGISwapChain1::Present1`, or legacy `IDXGISwapChain::Present` (`D2DERR_RECREATE_TARGET`, `DXGI_ERROR_DEVICE_REMOVED`, `DXGI_ERROR_DEVICE_RESET`, `DXGI_ERROR_DEVICE_HUNG`) MUST discard the full D3D/D2D device resource graph, not only the swap chain. FolderView must then mark the next paint as a full-client repaint and invalidate the pane so the following render rebuilds device resources from scratch.
- FolderView MUST use the shared DxUi `IsDeviceLossHResult(...)` predicate for
  render/present failure classification and the shared
  `CreateD3D11DeviceWithWarpFallback(...)` helper for D3D11 hardware-to-WARP
  fallback. It must not carry a private device-loss predicate or duplicate the
  D3D11 creation fallback sequence.
- Device-loss recovery MUST emit `folder.render.device_loss_recovery_count` with the failed operation as detail and the failing HRESULT on the metric row.
- Non-device-loss render/present failures keep the narrower swap-chain recreation path unless a later measurement or correctness case proves that broader discard is required.
- Out of memory: Reduce icon cache size, show error
- Transient Direct2D/DXGI/D3D failures during DPI changes, monitor scale changes, swap-chain resize, present, or device-target recreation MUST NOT create an end-user alert immediately. FolderView records `folder.render.failure_suppressed` diagnostics and keeps retry/recovery paths active.
- A user-visible **Rendering error** overlay is reserved for persistent rendering failures: the same rendering HRESULT must continue for at least three consecutive rendering-failure reports and at least 2 seconds without an intervening successful present/reset before promotion. Promotion emits `folder.render.failure_promoted`.
- Successful present, explicit rendering-overlay clear, or rendering-alert dismissal resets the rendering-failure persistence gate.
- Persistent swap chain recreation failure: Show error overlay and retry recreation (no GDI fallback; no `MessageBoxW`)

**Pattern:** report failures via the in-window overlay system (no `MessageBox*`).

Example: `FolderView::ReportError(L"EnumerateFolder", hr)` logs the failure and updates the current in-window overlay state.

## Make File List output

- File output from **Make File List** MUST use `Common::Files::LocalFileTransaction`; it must write the complete
  rendered UTF-8 document to a unique sibling, flush and size-check it, then replace the selected local target.
- Conversion, write, flush, verification, or promotion failure MUST preserve any previous target and remove the
  abandoned temporary sibling. Direct `CREATE_ALWAYS` writes to the requested final path are forbidden.
- The output transaction is only the commit boundary. Enumeration/rendering responsiveness, cancellation,
  progress, explicit destination selection, and overwrite consent remain command-workflow responsibilities.
- For interactive file output, the options prompt MUST be followed by a Save picker. The picker owns the final
  destination decision and MUST request overwrite consent before an existing target can be replaced; the raw
  remembered path is only a suggestion for that picker. Test automation may supply the destination directly.
- Local enumeration, sorting, text/CSV/JSON rendering, and file output MUST run on the pane's owned background
  `std::jthread`. The UI thread may snapshot command inputs and create/update the informational File Operations
  task, but it MUST return before recursive enumeration starts.
- Progress MUST remain visible through an informational File Operations task during collection, rendering, and
  output. Invoking Make File List again for the same pane requests cancellation; window teardown also requests
  stop and joins the worker. Collection and rendering MUST poll the stop token, and file output MUST recheck it
  after the temporary write and before promotion so cancellation cannot publish a new target.
- Clipboard publication remains a UI-thread operation. The worker may build clipboard text, but ownership crosses
  to the UI through `PostMessagePayload(...)`, and settings/success feedback update only after clipboard or file
  publication actually succeeds.

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
- `folderView_perf_sort_toggle_stress` is the focused sort baseline: it uses a 5,000-entry adversarial local folder, repeatedly toggles Name, Extension, Time, Size, and None sort modes, writes `folderView_perf_sort_toggle_stress_metrics.json`, emits `folder.sort_toggle_us`, and keeps `FolderView.ApplyCurrentSort` / `FolderView.ExecuteEnumeration.SortMerge` as the product metric families. This case is a deterministic metric recorder and structural regression guard, not a wall-clock CI threshold gate.
- FolderView sorting keeps medium interactive folders on the sequential `std::sort` path; parallel sort is reserved for genuinely large folders where scheduling overhead is amortized. `FolderViewSortPolicy::ShouldUseParallelSort(...)` defines and unit-tests the threshold boundary. The comparator must remain a total ordering by ending ties with `unsortedOrder`, which is what makes non-stable sorting behavior-preserving.
- Normal FolderView rendering must not apply or clear incremental-search text drawing effects unless incremental quick search is active with a non-empty query. Exiting quick search clears any existing DirectWrite layout effects once so inactive renders cannot leave stale highlighted text behind. `folderView_perf_sort_toggle_stress` guards inactive quick-search renders with `incrementalSearchEffectUpdates == 0`, and `cmd_pane_quickSearch_integrated_navigation` remains the active quick-search correctness guard.
- `folderView_dpi_change_repaints_both_panes` is the focused DPI/scale-transition guard: it warms both pane render targets, applies a synthetic per-monitor DPI change, verifies both panes update DPI before or during the transition, and verifies each pane schedules or completes a full-client render so no pane can expose a black/stale back buffer. 2026-06-09 Debug evidence is archived at `Specs/TestRuns/7d3a1247382a/Commands/2026-06-09_134243/`.
- `folderView_rendering_error_overlay_requires_persistence` is the focused rendering-error UX guard: it forces consecutive synthetic `IDXGISwapChain::Present` failures, verifies transient failures do not create a FolderView alert, verifies an aged third failure promotes to a rendering alert, verifies explicit clear hides the alert, and verifies the next single failure is suppressed again. The case also records `folder.render.failure_suppressed` and `folder.render.failure_promoted`; 2026-06-09 Debug evidence is archived at `Specs/TestRuns/7d3a1247382a/Commands/2026-06-09_134254/`.
- `folderView_render_device_loss_recovers` is the focused device-loss recovery guard: it injects one `DXGI_ERROR_DEVICE_REMOVED` at `EndDraw` and one at `Present`, verifies each path discards device resources, verifies the pane completes a full-client repaint with a valid D2D target afterward, verifies no rendering alert remains visible, and records `folder.render.device_loss_recovery_count`.
- `folderView_thumbnail_cached_only_no_close_stall` is the focused thumbnail fast-first, provider-handoff, and teardown-safety guard: it verifies the normal visible thumbnail path records zero provider-allowed shell lookups, drives one provider completion at the `50 ms` waiter boundary and requires exactly one `thumbnails.provider_probe_result_claimed` row, forces a later delayed provider probe, verifies navigation/close returns within the bounded deadline, verifies pending thumbnail work returns to zero, and records `thumbnails.close_to_idle_us`. Current-batch stale-generation bitmap payloads own and decrement their pending count; stale-batch and explicitly unaccounted late-provider payloads do not.
- `cmd_pane_clipboardPasteShortcut_returns_before_worker_complete`, `cmd_pane_clipboardPasteShortcut_close_does_not_wait_for_worker`, `cmd_pane_clipboardPasteShortcut_concurrent_invocations_create_distinct_links`, and `cmd_pane_clipboardPasteShortcut_failure_after_navigate_shows_alert` are the focused Paste Shortcut guards: they delay the shared `PasteShortcutSave` or `PasteShortcutAfterSlotProbe` hooks, verify command/navigation latency stays below the bounded deadline, force and recover one completion-post loss without wedging the queued invocation, verify concurrent invocations create distinct `.lnk` names, switch provider mid-work and prove the old local cache is invalidated without stale navigation/focus, verify forced create failures still surface after navigation away, and keep `cmd_pane_clipboardPasteShortcut_creates_unique_links` as the focus-on-last-shortcut regression guard.
- `folderView_perf_scroll_render_stress` is the focused normal-mode scroll/render baseline: it uses a 1,600-item folder in Brief, Detailed, and Extra Detailed modes, drives scroll through window messages, writes `folderView_perf_scroll_render_stress_metrics.json`, emits `folder.scroll_input_to_paint_us`, `folder.scroll_frame_count`, and `folder.scroll_visible_item_count`, and keeps `render.frame_us`, `render.layout_items_us`, and `render.draw_item_us` as compatibility product metric families. It must also preserve the frame-level metric contract for `folder.frame.total_us`, `folder.frame.present_us`, `folder.frame.visible_work_count`, `folder.frame.input_to_paint_us`, and `folder.frame.dirty_rect_area_px`.
- `DrawItem` must reuse cached member brushes for normal selected and hovered backgrounds. FolderView solid-brush creation routes through `CreateFolderViewSolidColorBrush(...)`, whose cached/transient lifetime classification increments the `ENABLE_TESTS` draw-item counter only when a real transient brush is created during item draw. `folderView_draw_item_brush_reuse_guard` requires zero for selected/hovered redraw and archives `folder.draw_item.transient_brush_create_count`; `folderView_perf_huge_folder_scale` keeps the same counter at zero while select-all scrolling a 10,000-item synthetic folder. Do not add a production or committed test-only force-allocation path inside `DrawItem`.
- The scroll/render baseline also records the real product `WM_PAINT` delta before deterministic warm-render frames through `folder.scroll.product_paint_render_count`, `folder.scroll.product_paint_full_client_count`, and `folder.scroll.product_paint_dirty_rect_area_px`, plus matching fields in `folderView_perf_scroll_render_stress_metrics.json`. Repeated boundary scroll requests that do not change the viewport MUST produce zero product paint frames. Viewport-changing scrolls currently repaint the full client; do not replace them with a partial dirty rect unless a separate DXGI scroll-rect path is gated and proves overlay, hover, selection, focus, resize, DPI, and device-loss correctness.
- DirectWrite item text-layout creation (label/details/metadata) is instrumented through the single seam `FolderView::CreateInstrumentedItemTextLayout`, which all three item-layout creation paths (`UpdateItemTextLayouts`, `EnsureItemTextLayout`, `ProcessIdleLayoutBatch`) route through. It emits the `dwrite.text_layout.create_count`, `dwrite.text_layout.create_us`, `dwrite.text_layout.frame_create_count`, and `dwrite.text_layout.frame_create_us` metric families, gated on `Debug::Perf::IsCaptureEnabled()` so normal builds pay nothing; `folderView_perf_scroll_render_stress` asserts their presence. 2026-06-19 same-machine evidence (`Specs/TestRuns/7d3a1247382a/Commands/`, six runs) found text-layout creation is **not** a material cost — about 1.2–1.4% of `render.layout_items_us` and present in only ~5 of ~108 render frames — so a DirectWrite text-layout cache is a measured no-op and MUST NOT be added without new evidence; the seam itself is where such a cache would hook if evidence ever changes. The dominant FolderView layout-pass cost is `render.layout_items_us` (p95 ~200ms versus frame p95 ~50ms), which is the open follow-up lead, not text-layout creation.
- `FolderView::LayoutItems` is decomposed into per-phase timing metrics `folder.layout.setup_us`, `folder.layout.estimate_metrics_us`, `folder.layout.column_resolve_us`, `folder.layout.bounds_us`, and `folder.layout.update_text_layouts_us` (gated on `Debug::Perf::IsCaptureEnabled()`), plus the `folder.layout.metrics_estimate_pass_count` and `folder.layout.update_visible_item_count` counters; `folderView_perf_scroll_render_stress` asserts the five `_us` keys are present. **Correction (2026-06-19):** an initial decomposition appeared to show `UpdateItemTextLayouts` owning ~82.6% of layout-pass time, but follow-up sub-decomposition found that was ~96% a *measurement artifact* — the perf JSONL sink opened and closed the file on every metric row (~50–150µs each), and the per-creation text-layout emits triggered ~3,748 such writes per run inside that function. After the sink was fixed to keep the append handle open (`Common/Common/PerfJsonl.cpp`), the layout pass is ~15ms p95 and `UpdateItemTextLayouts` ~9ms p95 (Debug) — **not** a bottleneck; the real per-item DirectWrite work (`GetMetrics` 2.2%, `SetMaxWidth` 0%, creation 1.4%) is negligible, so no layout optimization was warranted. See `Specs/Plans/Done/FolderView_LayoutPassDecomposition_MetricPilot_2026-06-19.md` and `Specs/Plans/Done/FolderView_UpdateItemTextLayouts_Optimization_2026-06-19.md`.
- `folderView_perf_overlay_invalidation_stress` is the focused overlay/quick-search metric guard: it opens a deterministic 180-item local folder, drives incremental search for `overlay`, shows the busy cancel overlay, lets the initial show animation settle, paints bounded overlay frames through the real FolderView render path with a limited message pump, writes `folderView_perf_overlay_invalidation_stress_metrics.json`, and records scenario-local presence counts for `folder.frame.overlay_animation_count`, `folder.frame.overlay_dirty_rect_area_px`, and `render.incremental_search_effect_updates`. The metric scan must begin at the case's initial `perf_metrics.jsonl` byte offset so earlier rows cannot satisfy the case. Overlay dirty-rect optimization claims must cite this scenario plus a same-machine non-overlay scroll comparison.
- Current FolderView performance proof is owned by the focused cases above. A regression or improvement claim must archive a fresh same-machine run for the affected case and comparison; unavailable historical run directories are not current regression anchors.
- `folderView_perf_directory_change_storm` is the pane-visible DirectoryInfoCache notification baseline: it opens a local folder in FolderView, applies deterministic create/rename/delete/directory churn, verifies final visible item count and focus stability, writes `folderView_perf_directory_change_storm_metrics.json`, and emits `folder.directory_change_storm_*`, `directorycache.post_refresh_count`, and the aggregate `folder.refresh.*` rows.
- `folderView_perf_refresh_preservation` is the focused same-folder refresh contract guard: it opens a local folder in FolderView, preserves focus and an active incremental-search query, mutates one created file, one deleted file, one selected same-folder rename, and a bounded create/delete burst, then verifies unchanged selection preservation, selected-rename transfer, focus survival, query survival, and one archived row for every `folder.refresh.*` metric.
- `folderView_refresh_to_paint_metric_clears_after_failed_render` is the focused failed-frame metric guard: it arms input-to-paint through real keyboard focus navigation and refresh-to-paint through a completed refresh, forces synthetic non-device-loss `EndDraw` failures, verifies neither pending sample emits during the failed frame, then drives later unrelated successful Presents and verifies neither stale sample leaks into them. Render-target absence plus `EndDraw`, `Present1`, and legacy `Present` failures all clear both pending metric slots.
- `folderView_perf_iconcache_contention` is the IconCache lock-diagnostic baseline: it opens dual icon-heavy panes with repeated unique extensions, writes `folderView_perf_iconcache_contention_metrics.json`, and reviews `iconcache.lock_wait_slow_us` / `iconcache.lock_hold_slow_us`. Slow-lock metric emission must run outside the measured IconCache mutex critical section; otherwise the diagnostic path can create the wait contention it reports. Treat hold-only rows as monitoring evidence; additional contention optimization requires repeated slow wait rows or an explicitly accepted hold-time bottleneck.
- `folderView_iconcache_live_path_failure_uses_bounded_backoff` is the focused live IconCache transient-failure guard: it forces two shell failures, verifies immediate repeats inside the first and doubled windows are negative-cache hits, verifies retries occur after each active window, and verifies the eventual success resets the negative state and is served from the positive cache thereafter. Its same-run count baseline/candidate rows are `icons.live_path_negative_backoff_shell_lookups.baseline_no_cache` and `.candidate`.
- `folderView_perf_icon_pipeline_cold_slow` is the focused icon-pipeline cold/slow guard: it clears IconCache, opens local per-file icon types (`.exe`, `.dll`, `.ico`, `.lnk`, `.url`, `.cpl`, `.scr`, `.msc`, `.ocx`), injects one slow HICON extraction through `SelfTestLatency::Point::IconExtractSystemIcon`, verifies at least one visible bitmap icon resolves before the delayed extraction finishes, opens a dummy offline/recall fixture, verifies per-file icon-index lookup consumes zero `IconPathLiveLookup` hooks, verifies `icons.recall_avoided_count`, and verifies synthetic offline/recall thumbnails fall back without provider-allowed shell I/O.
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

Optional view modes, grouping, marquee selection, inline rename, Quick Look,
resizable detail columns, and multi-key sorting are not current FolderView
contracts. Product decisions for those proposals are tracked only in
`Specs/Plans/WIP/Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md`.
Pane-to-pane copy and move are current behavior and are specified by
`Specs/FileSystem/FileSystem_FileOperations.md`.
