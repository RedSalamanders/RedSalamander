# Preferences Dialog Completeness Audit

**Author:** Ripley (Lead / Reviewer)
**Date:** 2025-07-25
**Scope:** All 11 Preferences panes — DxUI control wiring, settings round-trip, behavioral completeness
**Closeout audit:** 2026-04-25 - Done. The two historical code bugs called out below are now fixed in `Preferences.Dialog.cpp` and `Preferences.Viewers.cpp`; no remaining implementation gap was found in this audit.

---

## Per-Pane Audit

### 1. General Pane
**Status:** 🟢 Complete
**Controls:** 3 Toggles + 6 Labels + 3 CardPanels = 12 DxUI elements
**Settings wired:** 3/3 properly read+write

| Setting | Control | Read | Write | Sync | Callback |
|---------|---------|------|-------|------|----------|
| `mainMenu.menuBarVisible` | Toggle | ✅ | ✅ | ✅ | ✅ SetOnToggled |
| `mainMenu.functionBarVisible` | Toggle | ✅ | ✅ | ✅ | ✅ SetOnToggled |
| `startup.showSplash` | Toggle | ✅ | ✅ | ✅ | ✅ SetOnToggled |

**Missing:** None (spec mentions cache limits as "if appropriate" — not implemented, acceptable)
**Broken:** None
**Notes:** Clean implementation. Proper Settings-style cards with right-aligned toggles.

---

### 2. Panes Pane
**Status:** 🟢 Complete
**Controls:** 8 Toggles, 6 ComboBoxes, 1 TextField, 3 Headers, labels = ~30 DxUI elements
**Settings wired:** 11/11 properly read+write

| Setting | Control | Read | Write | Sync | Callback |
|---------|---------|------|-------|------|----------|
| Left `view.display` | Toggle + ComboBox | ✅ | ✅ | ✅ | ✅ |
| Left `view.sortBy` | ComboBox | ✅ | ✅ | ✅ | ✅ |
| Left `view.sortDirection` | Toggle + ComboBox | ✅ | ✅ | ✅ | ✅ |
| Left `view.statusBarVisible` | Toggle | ✅ | ✅ | ✅ | ✅ |
| Right `view.display` | Toggle + ComboBox | ✅ | ✅ | ✅ | ✅ |
| Right `view.sortBy` | ComboBox | ✅ | ✅ | ✅ | ✅ |
| Right `view.sortDirection` | Toggle + ComboBox | ✅ | ✅ | ✅ | ✅ |
| Right `view.statusBarVisible` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `folders.showHiddenFiles` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `folders.showSystemFiles` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `folders.historyMax` | TextField | ✅ | ✅ | ✅ | ✅ (digits-only, 1-50 range) |

**Missing:** None
**Broken:** None
**Notes:** High-contrast mode switches Display/SortDirection from Toggle to ComboBox — correct accessibility pattern. SortBy change auto-resets sortDirection to appropriate default. Six sync flags prevent callback re-entrancy.

---

### 3. Viewers Pane
**Status:** 🟢 Complete
**Controls:** 1 Grid, 2 TextFields, 1 ComboBox, 3 Buttons, 4 Labels, 1 Hint = 12 DxUI elements
**Settings wired:** 1/1 (`extensions.openWithViewerByExtension`) properly read+write

| Setting | Control | Read | Write | Sync | Callback |
|---------|---------|------|-------|------|----------|
| `extensions.openWithViewerByExtension` | Grid + ComboBox + Buttons | ✅ | ✅ | ✅ | ✅ Add/Update/Remove/Reset |

**Missing:** None per spec (search, list, add/update, remove, reset-to-defaults all present)
**Broken:** None. The prior `SyncDxEditsFromState()` dead-code issue is fixed; the function now syncs both editor fields from retained state during refresh.

**Notes:** Grid selection → editor fields → save button flow works correctly for normal usage. The dead sync is a correctness risk only in edge cases (e.g., external settings reload).

---

### 4. Plugins Pane
**Status:** 🟢 Complete
**Controls:** 2 Grids, 6 Buttons, 5 Labels = 14 DxUI elements
**Settings wired:** 3/3 properly read+write

| Setting | Control | Read | Write | Sync | Callback |
|---------|---------|------|-------|------|----------|
| `plugins.disabledPluginIds` | Grid checkbox | ✅ | ✅ | ✅ | ✅ OnGridCheckboxToggled |
| `plugins.customPluginPaths` | Grid + Add/Remove buttons | ✅ | ✅ | ✅ | ✅ Add/Remove |
| `plugins.configurationByPluginId` | Configure dialog | ✅ | ✅ | ✅ | ✅ Configure button |

**Missing:** None per spec
**Broken:** None
**Notes:** Correctly prevents disabling the active filesystem plugin (shows alert). Test/TestAll functionality works. Per-plugin child node navigation with configuration subpage. Search filtering. Very complete.

---

### 5. Compare Directories Pane
**Status:** 🟢 Complete
**Controls:** 11 Toggles, 1 ComboBox, 2 TextFields, 5 Headers = ~19+ DxUI elements
**Settings wired:** 14/14 properly read+write

| Setting | Control | Read | Write | Sync | Callback |
|---------|---------|------|-------|------|----------|
| `compareSize` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `compareDateTime` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `compareAttributes` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `compareContent` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `contentCompareWorkerCount` | ComboBox | ✅ | ✅ | ✅ | ✅ (0=Auto, 1-4) |
| `compareSubdirectories` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `compareSubdirectoryAttributes` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `selectSubdirsOnlyInOnePane` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `keepIdenticalItems` | Toggle | ✅ | ✅ | ✅ | ✅ (interdependent) |
| `showIdenticalItems` | Toggle | ✅ | ✅ | ✅ | ✅ (interdependent) |
| `ignoreFiles` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `ignoreFilesPatterns` | TextField | ✅ | ✅ | ✅ | ✅ (OnTextChanged + OnBlur) |
| `ignoreDirectories` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `ignoreDirectoriesPatterns` | TextField | ✅ | ✅ | ✅ | ✅ (OnTextChanged + OnBlur) |

**Missing:** None
**Broken:** None
**Notes:** Excellent interdependency logic: keepIdentical OFF → showIdentical forced OFF; showIdentical ON → keepIdentical forced ON. Ignore TextFields enabled only when corresponding toggle is ON. Layout reconstructs on toggle changes.

---

### 6. Hot Paths Pane
**Status:** 🟢 Complete
**Controls:** Per slot ×10: PathTextField, BrowseButton, LabelTextField, ShowInMenuToggle + headers/labels. Plus global toggle = ~41 interactive DxUI elements
**Settings wired:** 31/31 properly read+write

| Setting | Control | Read | Write | Sync | Callback |
|---------|---------|------|-------|------|----------|
| `hotPaths.slots[0..9].path` ×10 | TextField | ✅ | ✅ | ✅ | ✅ OnTextChanged + OnBlur |
| `hotPaths.slots[0..9].label` ×10 | TextField | ✅ | ✅ | ✅ | ✅ OnTextChanged + OnBlur |
| `hotPaths.slots[0..9].showInMenu` ×10 | Toggle | ✅ | ✅ | ✅ | ✅ SetOnToggled |
| `hotPaths.openPrefsOnAssign` | Toggle | ✅ | ✅ | ✅ | ✅ SetOnToggled |

**Missing:** None per spec
**Broken:** None
**Notes:** Label/ShowInMenu disabled when path is empty — correct UX. IFileOpenDialog with FOS_PICKFOLDERS for Browse. Whitespace trimming on fields. Slot auto-created when first value set.

---

### 7. Advanced Pane
**Status:** 🟢 Complete (for exposed settings)
**Controls:** 14 Toggles, 6 TextFields, 1 ComboBox, 4 Headers = ~25 interactive DxUI elements
**Settings wired:** 21/21 exposed settings properly read+write

| Setting | Control | Read | Write | Sync | Callback |
|---------|---------|------|-------|------|----------|
| `connections.bypassWindowsHello` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `connections.allowInsecureTlsInAutomation` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `connections.windowsHelloReauthTimeoutMinute` | TextField | ✅ | ✅ | ✅ | ✅ (digits, validated) |
| `monitor.menu.toolbarVisible` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `monitor.menu.lineNumbersVisible` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `monitor.menu.alwaysOnTop` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `monitor.menu.showIds` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `monitor.menu.autoScroll` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `monitor.filter.preset` | ComboBox | ✅ | ✅ | ✅ | ✅ |
| `monitor.filter.mask` (5 bit toggles) | 5 Toggles | ✅ | ✅ | ✅ | ✅ |
| `monitor.filter.mask` (direct edit) | TextField | ✅ | ✅ | ✅ | ✅ (0-31) |
| `cache.directoryInfo.maxBytes` | TextField | ✅ | ✅ | ✅ | ✅ (size suffix: MB/GB) |
| `cache.directoryInfo.maxWatchers` | TextField | ✅ | ✅ | ✅ | ✅ (optional) |
| `cache.directoryInfo.mruWatched` | TextField | ✅ | ✅ | ✅ | ✅ (optional) |
| `fileOperations.diagnosticsInfoEnabled` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `fileOperations.diagnosticsDebugEnabled` | Toggle | ✅ | ✅ | ✅ | ✅ |
| `fileOperations.maxDiagnosticsLogFiles` | TextField | ✅ | ✅ | ✅ | ✅ (digits, non-zero) |

**Settings in schema WITHOUT UI (intentional, expert-level):**
- `fileOperations.autoDismissSuccess` — managed by file operations dialog
- `fileOperations.maxIssueReportFiles` — deep internal
- `fileOperations.maxDiagnosticsInMemory` — deep internal
- `fileOperations.maxDiagnosticsPerFlush` — deep internal
- `fileOperations.diagnosticsFlushIntervalMs` — deep internal
- `fileOperations.diagnosticsCleanupIntervalMs` — deep internal

**Missing:** None (the 6 unexposed settings are appropriately JSON-only)
**Broken:** None
**Notes:** Filter mask toggles disabled unless preset is "Custom" — good UX. Cache bytes accepts size suffixes. All TextFields have proper validation and sync flags.

---

### 8. Keyboard Pane
**Status:** 🟢 Complete
**Controls:** 1 Grid, 1 TextField (search), 1 ComboBox (scope), 5 Buttons, 3 Labels = 11 DxUI elements
**Settings wired:** 1/1 logical setting (`shortcuts.functionBar` + `shortcuts.folderView`)

| Setting | Control | Read | Write | Sync | Callback |
|---------|---------|------|-------|------|----------|
| `shortcuts.*` | Grid + Assign/Remove/Reset + Import/Export | ✅ | ✅ | ✅ | ✅ Full lifecycle |

**Missing:** None per spec
**Broken:** None
**Notes:** Full capture mode for shortcut assignment with live chord preview. Inline conflict detection without modal dialogs. Import/Export JSON. Scope filter (All/FunctionBar/FolderView). Search filtering. Controls disabled during capture. Very complete implementation.

---

### 9. Themes Pane
**Status:** 🟢 Complete
**Controls:** 1 Grid, 2 ComboBoxes, 4 TextFields, 1 ColorSwatch, 7 Buttons, 6 Labels = ~22 DxUI elements
**Settings wired:** 2/2 (`theme.currentThemeId`, `theme.themes`)

| Setting | Control | Read | Write | Sync | Callback |
|---------|---------|------|-------|------|----------|
| `theme.currentThemeId` | ComboBox | ✅ | ✅ | ✅ | ✅ SetOnSelectionChanged |
| `theme.themes` (user themes) | Full editor suite | ✅ | ✅ | ✅ | ✅ Multiple callbacks |

**Behavioral features:**
- Theme ComboBox: builtin + user + `<New Theme>` entry ✅
- New theme workflow: name + base theme ✅
- Colors grid: key/swatch/value display ✅
- Color editor: hex text + swatch preview ✅
- Pick Color button: opens color picker ✅
- Set/Remove override per key ✅
- Load From File / Duplicate / Save Theme ✅
- Apply Temporarily (transient preview, Cancel reverts) ✅
- Override indicator (editable vs read-only themes) ✅
- Search/filter colors ✅

**Missing:** None per spec
**Broken:** None
**Notes:** Full theme lifecycle implementation. Transient preview correctly reverts on Cancel via baseline settings restore. Builtin themes read-only, user themes fully editable. Auto-generates `user/<slug>` IDs from theme name.

---

### 10. Editors Pane
**Status:** 🔴 Placeholder (by design)
**Controls:** 0 interactive DxUI controls
**Settings wired:** 0/0

**Missing:** Everything — spec explicitly states "Placeholder page (data model TBD)"
**Broken:** N/A
**Notes:** Displays `IDS_PREFS_EDITORS_PLACEHOLDER` static text only. The underlying connection settings (`connections.items`) exist and are managed by dedicated connection dialogs outside Preferences. This is the correct approach per spec.

---

### 11. Mouse Pane
**Status:** 🔴 Placeholder (by design)
**Controls:** 0 interactive DxUI controls
**Settings wired:** 0/0

**Missing:** Everything — spec explicitly states "Placeholder page (data model TBD)"
**Broken:** N/A
**Notes:** Displays `IDS_PREFS_MOUSE_PLACEHOLDER` static text only. Awaiting action/shortcut model finalization per spec.

---

## Dialog Infrastructure (Preferences.Dialog.cpp)

| Feature | Status | Notes |
|---------|--------|-------|
| Apply button | ✅ Working | CommitAndApply → Save → post kSettingsApplied |
| OK button | ✅ Working | Apply if dirty, then close |
| Cancel button | ✅ Working | Reverts transient theme preview, closes |
| Dirty tracking | ✅ Working | Exhaustive comparison across all settings domains |
| External reload detection | ✅ Working | staleFromExternalReload flag + user prompt |
| Theme refresh pipeline | ✅ Working | Full palette update after Apply |
| Global Reset to Defaults | ❌ Missing | No global "Reset All" button (per-pane resets exist for Viewers, Keyboard) |
| Lazy pane creation | ✅ Working | Panes created on first selection |
| Scroll retention | ✅ Working | Per-pane scroll position saved in retainedPageScrollYByCategory[] |

---

## Resolved Bug: `SaveSettingsFromDialog` Missing `showHiddenFiles`/`showSystemFiles` Merge

**Severity:** Medium — data loss on save in specific scenarios
**Location:** `Preferences.Dialog.cpp` lines 1902-1927

**The problem:** `SaveSettingsFromDialog()` creates a `merged` settings object by first snapshotting the owner's current runtime state (via `kPreferencesRequestSettingsSnapshot`), then selectively overwriting changed fields from `workingSettings`. For folder settings, it explicitly merges:
- `historyMax` ✅
- Per-pane `view.display`, `view.sortBy`, `view.sortDirection`, `view.statusBarVisible` ✅

But it does **NOT** explicitly merge:
- `folders.showHiddenFiles` ❌
- `folders.showSystemFiles` ❌

The `AreEquivalentFolderPreferences()` function correctly checks these fields (line 1485), so dirty detection works. But the save merge path doesn't include them.

**Impact:** If the owner's snapshot response overwrites `*state.settings` with runtime state that has the old `showHiddenFiles`/`showSystemFiles` values, the merged object will contain the old values, and the save to disk will lose the user's changes. The in-memory state is fixed afterward (line 2112: `*state.settings = state.workingSettings`), so the change takes effect at runtime but won't survive an app restart.

**Resolution:** Fixed. `SaveSettingsFromDialog()` now snapshots and writes `folders->showHiddenFiles` and `folders->showSystemFiles` in the explicit folder merge block alongside `historyMax`.

---

## Resolved Bug: Viewers `SyncDxEditsFromState` Dead Code

**Severity:** Low — cosmetic, edge-case staleness
**Location:** `Preferences.Viewers.cpp` line 536

**The problem:** `SyncDxEditsFromState()` begins with `if (true) { return; }`, making the extension and viewer editor TextFields never sync from state. They're only populated via `UpdateEditorFromSelection()` on grid selection change.

**Impact:** If an external event (settings reload, theme change) triggers a full refresh without re-selecting the grid row, the editor fields could show stale data. Normal usage flow (user selects row → edits → saves) works correctly.

**Resolution:** Fixed. `SyncDxEditsFromState()` now updates the search and extension editor fields while guarding callback re-entrancy with `_syncingDxEdits`.

---

## Settings Coverage Summary

### Settings WITH Preferences UI: 80+ individual settings across 9 active panes

| Domain | Settings | UI Controls | Status |
|--------|----------|-------------|--------|
| General (mainMenu, startup) | 3 | 3 Toggles | 🟢 100% |
| Panes (folders view) | 11 | 8 Toggles + 6 Combos + 1 TextField | 🟢 100% |
| Viewers (extensions) | 1 map | Grid + ComboBox + Buttons | 🟢 100% |
| Plugins | 3 | 2 Grids + 6 Buttons | 🟢 100% |
| Compare Directories | 14 | 11 Toggles + 1 Combo + 2 TextFields | 🟢 100% |
| Hot Paths | 31 | 10×(TextField+Button+TextField+Toggle) + Toggle | 🟢 100% |
| Advanced | 21 exposed | 14 Toggles + 6 TextFields + 1 Combo | 🟢 100% |
| Keyboard | 1 logical | Grid + 5 Buttons + Search + Scope | 🟢 100% |
| Themes | 2 + colors | Grid + 2 Combos + 4 TextFields + 7 Buttons + Swatch | 🟢 100% |
| Editors | 0 | Placeholder | 🔴 By design |
| Mouse | 0 | Placeholder | 🔴 By design |

### Settings WITHOUT UI (by design — not bugs):

| Setting | Reason |
|---------|--------|
| `folders.active`, `folders.layout.*`, `folders.history*`, `folders.items` | Runtime state managed by app |
| `search.*` (16 settings) | Managed by Search dialog |
| `selectionMasks.*` (3 settings) | Runtime MRU state |
| `windows.*` | Window placement (runtime) |
| `connections.items` | Managed by connection dialogs |
| `plugins.currentFileSystemPluginId` | Managed by plugin system |
| `extensions.openWithFileSystemByExtension` | No UI for FS extension mapping (planned?) |
| `fileOperations` internals (6 settings) | Expert-level, JSON-only |

---

## Prioritized Fix List

### Closed
1. **Fix `SaveSettingsFromDialog` showHiddenFiles/showSystemFiles merge** — fixed before closeout.
2. **Remove dead `if (true) return;` in Viewers `SyncDxEditsFromState()`** — fixed before closeout.

### P2 — Feature Completeness (future)
1. **Editors pane implementation** — blocked on data model (connections vs. external editors)
2. **Mouse pane implementation** — blocked on action/shortcut model
3. **Global "Reset All to Defaults" button** — currently only per-pane resets exist
4. **`extensions.openWithFileSystemByExtension` UI** — no UI for filesystem plugin extension mapping

---

## Overall Assessment

**The Preferences dialog is in excellent shape.** 9 of 11 panes are fully wired and functional, covering 80+ settings with proper read/write/sync cycles. The 2 placeholder panes are explicitly spec'd as placeholders. The former showHiddenFiles/save merge bug and Viewers sync dead code issue are both resolved, so this audit is complete.

Architecture quality is high: consistent callback patterns, sync flags preventing re-entrancy, proper dirty tracking, theme refresh pipeline, external reload conflict detection, and lazy pane creation. The DxUI migration is essentially complete for all functional panes.
