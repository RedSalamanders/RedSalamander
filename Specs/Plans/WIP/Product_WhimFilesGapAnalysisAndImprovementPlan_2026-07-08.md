# RedSalamander vs. the File-Manager Field — Feature Gap Analysis & Improvement Plan

**Date:** 2026-07-08
**Status:** Draft for review (no code changes yet)
**Author:** Competitive analysis pass
**Competitors analysed:** WhimFiles · Total Commander · File Pilot · Directory Opus · XYplorer · Files (files.community) · fman · and a broader field (Double Commander, FreeCommander, Multi Commander, Q-Dir, Far Manager, One Commander, Spacedrive, terminal managers).
**Scope:** Compare the full feature surface of RedSalamander against the leading dual-pane / power file managers, identify what RedSalamander is missing or does worse, and specify the improvements worth building.

**Document structure:**
- **Part I (§1–§7)** — WhimFiles (macOS, filtering-focused) + the core gap specs **G1–G11**.
- **Part II (§8–§12)** — Total Commander (genre benchmark) + File Pilot (performance next-gen) + the new gaps **G12–G18**, plus a consolidated roadmap.
- **Part III (§13–§19)** — the broader landscape, deep analysis of **Directory Opus** and **XYplorer** (the premium tier) plus Files/fman/the field, the new gaps **G19–G24** they expose, and the **final consolidated roadmap (supersedes §11)**.

Reference sources: `docs/UserGuide.md`, `Specs/UI/UI_FolderView.md`, `Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/UI/UI_NavigationView.md`; WhimFiles pages + MacUpdate; Total Commander ghisler.com feature pages + v11 notes; File Pilot filepilot.tech + reviews; Directory Opus gpsoft.com.au; XYplorer xyplorer.com; Files files.community; fman fman.io (all captured 2026-07-08).

---

## 1. Executive Summary

RedSalamander and WhimFiles are both dual-pane file managers, but they sit at opposite ends of the design spectrum:

- **RedSalamander** is a Windows-native *power tool*: remote/cloud file systems (FTP/SFTP/SCP/IMAP/S3/S3Table/Google Drive/OneDrive/SharePoint), a plugin architecture, nine viewer plugins, Compare Directories with sync, an indexed search service, an ETW monitor, and full theming. It is broad and deep.
- **WhimFiles** is a macOS-native *focused, delightful* tool: a 9 MB, $19.99 one-time app whose entire identity is **speed and filtering**. It does far less, but the handful of things it does — live multi-dimensional filtering, fuzzy navigation, tabs, a command palette, undo, instant preview — are extremely polished and are exactly the everyday interactions users touch most.

**The gap is not breadth — RedSalamander wins on breadth by a wide margin. The gap is in the lightweight, high-frequency "quality of life" interactions that WhimFiles has made its whole product.** Those are the features RedSalamander users would feel every single session.

The verified actionable gaps, in priority order:

| # | Gap | RedSalamander today | Priority |
|---|-----|---------------------|----------|
| G1 | **Live multi-dimensional filtering** (type + date + size at once) + preset/named type categories | Name/wildcard mask only (`NameFilterState{enabled,text}`) | **P0** |
| G2 | **Undo for file operations** (copy/move/rename/recycle) | None — app mutates files with no undo (only a TODO spike, `plans/013`) | **P0** |
| G3 | **Tabbed browsing** (multiple tabs per pane) | Strictly two panes, one folder each (`Pane{Left,Right}`) | **P1** |
| G4 | **True command palette** (fuzzy, all commands) | Partial — F1 "Display Shortcuts" is substring-only and lists bound commands only | **P1** |
| G5 | **Fuzzy Go to Folder / Go to File** quick-jump | None (only per-level address autocomplete + in-folder Quick Search) | **P1** |
| G6 | **File checksums** (SHA-256/MD5/CRC32) in Properties/command | None surfaced (Properties → shell; compare is byte-by-byte) | **P1** |
| G7 | **First-class HEIC/WebP/AVIF** + quick/batch image convert | Partial — WIC-latent formats not registered; single-image export only | **P2** |
| G8 | **Bookmarks/Favorites** beyond 10 slots + a Places sidebar | Partial — 10 Hot Paths in a dropdown, no sidebar/tree | **P2** |
| G9 | **Hover preview** (image/PDF peek without opening) | Partial — has a docked Preview Pane, not hover | **P2** |
| G10 | **Always-calculate folder sizes** column mode | Partial — on-demand (Space key) + ViewerSpace treemap, no auto column | **P2** |
| G11 | **"Open in app" terminal helper** (`rs`/`redsal` shell function) | Partial — has "open shell here", not the reverse convenience | **P3** |

Sections 3–4 give the full comparison and per-gap specifications for WhimFiles. **Part II (§8 onward) extends the analysis to Total Commander and File Pilot**, which reinforce several of these gaps (tabs, command palette, fuzzy navigation, filtering) and add seven more (G12–G18: customizable toolbar, flexible pane layout, directory tree, persistent command line, file-type colors, split/combine, encode + checksum files). **Part III (§13 onward) adds the premium tier — Directory Opus and XYplorer — plus the modern/open field**, exposing six deeper gaps (G19–G24: file tags/labels/ratings, user scripting, duplicate finder, rich configurable metadata columns, collections/saved searches, and a power-user long-tail bundle).

---

## 2. Product Positioning

| Dimension | RedSalamander | WhimFiles |
|-----------|---------------|-----------|
| Platform | Windows 11 (x64 + ARM64), native C++23 / Direct2D / DirectWrite | macOS 12+, Apple Silicon only, native |
| Footprint | Large multi-project solution + plugins | ~9 MB single app |
| Distribution | GitHub, MSIX/MSI, winget | Lemon Squeezy, 30-day trial |
| Pricing | Free / open | $19.99 one-time (launch), $24.99 regular |
| Philosophy | Broad, deep, extensible power tool | Narrow, fast, opinionated daily driver |
| Core differentiator | Remote/cloud FS + plugins + viewers + compare | Real-time filtering + speed + keyboard flow |

The takeaway: **RedSalamander should adopt WhimFiles' interaction polish without giving up its breadth.** None of the gaps below require abandoning anything RedSalamander already does well; they are additive.

---

## 3. Full Feature Comparison Matrix

Legend: ✅ full · 🟡 partial/weaker · ❌ absent · ➖ not applicable

| Feature area | RedSalamander | WhimFiles | Gap? |
|---|:--:|:--:|:--:|
| Dual-pane layout | ✅ Left/Right, maximize, swap, split ratio | ✅ dual-pane | — |
| **Tabbed browsing** | ❌ | ✅ ⌘T/⌘W/⌘1–9, drag-reorder, per-tab state | **G3** |
| View modes | ✅ Brief / Detailed / Extra / Thumbnails / Preview | ✅ list + icon preview | — |
| **Live filtering (type+date+size)** | 🟡 name/wildcard mask only | ✅ simultaneous type + date + size | **G1** |
| **Preset filter categories** (Images/Videos/Docs/PDF) | ❌ | ✅ presets + custom named types | **G1** |
| Recursive/flatten view | 🟡 via Find window | ✅ recursive filter toggle | G1 |
| In-folder incremental search | ✅ Quick Search (Shift+Space) | ✅ | — |
| Full search (name/content) | ✅ native + indexed + fallback + search service | 🟡 recursive filter only | RS wins |
| **Fuzzy Go to Folder / File** | ❌ | ✅ ⌘G / ⌘P | **G5** |
| **Command palette** | 🟡 F1 shortcuts window (substring, bound-only) | ✅ ⇧⌘A all actions | **G4** |
| Rich viewers (text/hex/img/RAW/SQLite/PE/media/web/json/md) | ✅ 9 plugins | ❌ (Quick Look only) | RS wins |
| **Hover preview** | 🟡 docked Preview Pane | ✅ hover image/PDF peek | **G9** |
| Batch rename | ✅ deep (macros, case, manual, validation, modeless) | ✅ find&replace, numbering, case | — |
| Copy/move/delete engine | ✅ background, pause/resume, speed limit, parallel, conflict prompts, cross-FS bridge | 🟡 cut/copy/paste + conflict handling | RS wins |
| **Undo file operations** | ❌ | ✅ moves + trash | **G2** |
| **Image format convert (HEIC/WebP/AVIF→JPG/PNG)** | 🟡 single-image viewer export (PNG/JPEG/TIFF/BMP/GIF/JXR) | ✅ right-click convert | **G7** |
| **File checksum (SHA-256/MD5/CRC32)** | ❌ | ✅ Get Info SHA-256 | **G6** |
| Archive as virtual FS | ✅ 7z browse + ZIP pack/unpack | 🟡 ZIP create only | RS wins |
| Remote / cloud file systems | ✅ FTP/SFTP/SCP/IMAP/S3/S3Table/GDrive/OneDrive/SharePoint | ❌ (local + mounted/network via Finder) | RS wins |
| Compare directories + sync | ✅ | ❌ | RS wins |
| **Bookmarks / favorites** | 🟡 10 Hot Paths (dropdown) | ✅ sidebar bookmarks (unlimited) + Locations | **G8** |
| **Places sidebar / tree** | ❌ (drive/menu dropdown) | ✅ sidebar w/ volumes, drives, shares, eject | **G8** |
| **Folder size auto-fill** | 🟡 on-demand + treemap (ViewerSpace) | ✅ "Always Calculate" column | **G10** |
| Themes | ✅ System/Light/Dark/Rainbow/HighContrast + user JSON5 | 🟡 follows macOS appearance | RS wins |
| Plugin extensibility | ✅ FS + viewer plugins | ❌ | RS wins |
| Diagnostics tooling | ✅ RedSalamanderMonitor (ETW) | ❌ | RS wins |
| **Open-in-app from terminal** | 🟡 "open shell here" (reverse) | ✅ `wf` shell function | **G11** |
| Session persistence | ✅ pane paths, view, sort, history, window | ✅ tabs, folders, filters | — |

**Net:** RedSalamander is ahead or equal on 13 rows and behind on the 11 gaps (G1–G11). Every gap is in the *daily-interaction* layer, which is why they matter disproportionately.

---

## 4. Gap Specifications

Each gap below follows the same structure: **Current state → Proposed design → Commands/settings → Acceptance criteria → Effort/risks.** Command IDs follow the existing `cmd/pane/*` and `cmd/app/*` convention (`Specs/UI/UI_CommandMenuKeyboard.md`). New selftests must join the Commands self-test suite (see `Tests/README.md` and the `implemented_menu_labels_not_todo` guard pattern).

---

### G1 — Live Multi-Dimensional Folder Filtering + Named Type Categories `[P0]`

**Why it matters most:** This is WhimFiles' entire reason to exist ("The file manager built around filtering"). It is the single highest-leverage feature RedSalamander can adopt, and it extends machinery that already exists.

**Current state**
- The pane filter (`cmd/pane/filter`, `Ctrl+F12`) and the Filter Bar are a **wildcard name mask only**:
  - `RedSalamander/FolderView.h:522` — `struct NameFilterState { bool enabled = false; std::wstring text; };`
  - `RedSalamander/FolderView.h:1127` — `struct CompiledNameFilter { NameFilterState state; MaskSyntax::WildcardMask mask; bool hasMask; };`
  - `RedSalamander/MaskSyntax.h:10` — mask carries only `includePatterns` / `excludePatterns` (glob strings).
  - `Specs/UI/UI_FolderView.md:311,503` — filter excludes items whose `displayName` does not match the mask.
- Size and date exist as dimensions **only** in Compare Directories (`RedSalamander.rc:617,619`), which compares two trees — it does not narrow a single folder's live view.
- No preset categories; the "Documents/Pictures/Videos" strings are Windows known-folder navigation links, not type filters.

**Proposed design**

1. **Structured filter criteria.** Replace the name-only filter state with a composable predicate applied during enumeration/display:

   ```cpp
   struct SizeRange   { std::optional<uint64_t> minBytes, maxBytes; };
   struct DateRange   { std::optional<FILETIME> from, to;           // absolute
                        std::optional<uint32_t> withinDays; };       // relative (today/7/30/365)
   struct FolderFilterCriteria {
       bool                       enabled = false;
       std::wstring               nameMask;            // existing wildcard include/exclude
       std::optional<std::wstring> typeGroupId;        // e.g. "images", "video", user group id
       std::optional<DateRange>   modified;
       std::optional<SizeRange>   size;
       bool                       recursive = false;   // flatten subfolders into the filtered view
       bool                       matchFoldersToo = false;
   };
   ```
   All present criteria are ANDed. Items are hidden unless every active dimension matches. The existing `CompiledNameFilter` becomes one term of the compiled predicate.

2. **File-type groups** (new settings model), with built-in presets plus user-defined named groups reusable across every folder:

   ```cpp
   struct FileTypeGroup { std::wstring id, displayName; std::vector<std::wstring> extensions; bool builtin; };
   struct FileTypeGroupsSettings { std::vector<FileTypeGroup> groups; };
   ```
   Built-in presets: **Images** (`.jpg .jpeg .png .gif .bmp .tif .tiff .webp .heic .avif .jxr …`), **Video**, **Audio**, **Documents**, **PDF**, **Archives**, **Code/Text**, **Executables**. Users add/edit groups in a new **Preferences → Filters** page (reuse the existing schema-driven Preferences editor).

3. **UI — upgrade the Filter Bar into a live filter row.** Below the pane, show removable "chips" for each active dimension (e.g. `Images ✕`, `> 100 MB ✕`, `Modified: last 7 days ✕`, `Recursive ✕`). A filter button opens a popup with: type-group dropdown, size min/max with quick presets (`>1 MB`, `>100 MB`, `<1 KB`), date pickers + relative quick buttons (Today / 7d / 30d / This year), and a Recursive toggle. Typing in the existing filter combo still edits `nameMask`. Filtering is incremental (updates as you type/toggle), matching WhimFiles' "instant" feel — reuse the async enumeration + generation-staleness pattern already in `FolderView.Enumeration.cpp`.

4. **Recursive mode** flattens descendants into the filtered list (a lightweight, filter-scoped version of Find), so "show all PDFs under this tree" is one toggle, not a separate window.

**Commands / settings**
- Extend `cmd/pane/filter` to open the new multi-dimension popup.
- Add `cmd/pane/filter/typeGroup/<groupId>` (one-click category), `cmd/pane/filter/recursive` (toggle), `cmd/pane/filter/clear`.
- New `settings.fileTypeGroups`; extend the per-pane filter persistence in `FoldersSettings`/`FolderViewSettings` (`Common/SettingsStore.h:99-143`) to store the full `FolderFilterCriteria` (currently only a name mask is persisted).
- Update `Specs/SettingsStore.schema.json` + `SettingsSchemaTests`.

**Acceptance criteria**
- With Images preset + `> 100 MB` + `modified within 30 days` all active, only items satisfying all three appear; removing any chip re-broadens the list immediately.
- Type groups persist and apply identically across different folders and across restart.
- Recursive toggle lists matching descendants; clearing it returns to the current folder only.
- Name mask, type group, size, and date compose (AND) and each is independently clearable.
- New Commands self-test `cmd_pane_filter_multidimension_*` (RED→GREEN) covering compose/clear/persist/recursive; `SettingsSchemaTests` covers the new schema.

**Effort:** L (largest of the P0s — new predicate engine, settings, Preferences page, Filter Bar UI). **Risk:** enumeration hot-path perf (mitigate with the existing async/staleness machinery and short-circuit predicate evaluation); recursive mode memory on huge trees (bound + cancel on navigate, reuse Find's guards).

---

### G2 — Undo for File Operations `[P0]`

**Why it matters:** RedSalamander mutates user files and offers no undo — a safety gap the project already recognizes (`plans/013-undo-file-operations-spike.md`, status TODO: *"RedSalamander mutates user files — and offers no undo"*). WhimFiles ships undo for moves and trash. Given RedSalamander's documented history of data-loss reviews (see memory: Riptide/Fairstream/Floodgate cloud data-safety findings), a robust undo is high-value insurance, not just polish.

**Current state**
- No `Undo`/`Redo` command in `CommandRegistry.cpp` or `ShortcutDefaults.cpp`.
- Only text-edit undo exists, and only inside DxUi text inputs (`Common/DxUi/DxUi.TextInput.cpp:933-969`) — unrelated to file operations.
- Batch Rename exports a "Copy Undo Plan" (a TSV report) described in its own docs as *"a manual escape hatch, not an undo."*

**Proposed design**

1. **Operation journal.** Each completed mutating operation records an inverse action on a bounded, session-scoped undo stack:

   | Operation | Inverse | Undoable? |
   |---|---|---|
   | Move (same volume) | Move back | ✅ |
   | Move (cross-volume/cross-FS) | Move back (re-transfer) | ✅ with confirm (cost warning) |
   | Rename | Rename back | ✅ |
   | Copy | Delete the created copies | ✅ |
   | MakeDir | Remove the created dir (if still empty) | ✅ |
   | Recycle-bin Delete | Restore from Recycle Bin (shell `IFileOperation` undo / restore) | ✅ |
   | **Permanent Delete** | — | ❌ (must stay non-undoable; warn before) |

2. **Recording point.** Hook the file-operation completion path (the engine behind `F5`/`F6`/`F8`, `FolderWindow.FileOperations.cpp`) to push a journal entry with the concrete source/destination set and the chosen conflict resolutions, so the inverse is exact.

3. **Undo/redo semantics.** `cmd/app/undo` (`Ctrl+Z`) pops and executes the inverse as a normal tracked operation (progress, conflict prompts, cancel). `cmd/app/redo` (`Ctrl+Y`) re-applies. Undo of a partially-completed operation only reverts what actually happened (use the engine's per-item success ledger). Cross-FS undo shows an explicit confirmation because it re-transfers bytes.

4. **Visibility.** After an operation, the status bar / File Operations popup shows a transient "Undo: Moved 3 items" affordance. Edit menu gains Undo/Redo with the last action described.

**Commands / settings**
- `cmd/app/undo`, `cmd/app/redo`; Edit-menu entries; defaults `Ctrl+Z` / `Ctrl+Y` (verify no conflict in `ShortcutDefaults.cpp`).
- Setting: undo stack depth (default e.g. 20) and an on/off toggle under Preferences → File Operations.
- Consume/close the `plans/013` spike; produce the design doc it references (`Specs/Plans/WIP/Core_FileOperationsUndo.md`, currently absent).

**Acceptance criteria**
- Move → Undo restores originals to source paths and removes them from destination.
- Rename → Undo restores the original names.
- Copy → Undo deletes exactly the created copies (never pre-existing files at the destination).
- Recycle delete → Undo restores from the Recycle Bin; Permanent delete offers **no** undo and warns beforehand.
- Undo of a cancelled/partial op reverts only completed items.
- Cross-volume/cross-FS undo prompts before re-transfer.
- New FileOps self-test family `FileOps_Undo_*` (join `kFileOpsFamilyDefinitions` or Full runs skip it — see memory `fileops-selftest-family-registration`) covering each row above, including the partial-op and permanent-delete-not-undoable cases.

**Effort:** M–L. **Risk:** correctness of inverse under conflicts/partial completion and cross-FS — must reuse the engine's existing per-item ledger and never guess. Keep permanent delete strictly non-undoable.

---

### G3 — Tabbed Browsing (Tabs per Pane) `[P1]`

**Current state**
- The pane model is a fixed two-value enum, one folder each: `RedSalamander/FolderWindow.h:304` `enum class Pane { Left, Right };`; each pane holds a single `current` path.
- A reusable `TabControl` already exists in the framework — `Common/DxUi/DxUi.h:1923` with `AddTab`/`CloseTab`/`ReorderTab` — but it is used only for the Preview Pane's Folder/Preview toggle (`FolderWindow.cpp:1793`) and Preferences dialog pages, **not** folder browsing.

**Proposed design**
- Give each pane an ordered collection of tabs plus an active index; each tab owns its own folder path, view mode, sort, filter criteria (G1), selection, and navigation history. The pane renders the active tab; the other panes/engine keep addressing "the focused pane" exactly as today.
- Reuse the existing DxUi `TabControl` for the tab strip (drag-reorder already supported). Show the strip only when a pane has >1 tab (or per a Preferences toggle), to keep the single-folder look for users who don't want tabs.
- Persist tabs by extending `FoldersSettings.items` (`Common/SettingsStore.h:112-143`) from one `FolderPane` per side to a list per side + active index. Session restore already rebuilds per-pane state (`RedSalamander.cpp:9174-9345`) — extend it to iterate tabs.

**Commands / keyboard**
- `cmd/pane/newTab`, `cmd/pane/closeTab`, `cmd/pane/duplicateTab`, `cmd/pane/nextTab`, `cmd/pane/prevTab`, `cmd/pane/reopenClosedTab`, `cmd/pane/gotoTab/<n>`.
- Suggested defaults: `Ctrl+T`, `Ctrl+W`, `Ctrl+Tab` / `Ctrl+Shift+Tab`, `Ctrl+Shift+T`.
- **Shortcut conflict to resolve:** WhimFiles uses `⌘1–9` for tab jump, but RedSalamander already binds `Ctrl+1…Ctrl+0` to Hot Paths (`ShortcutDefaults.cpp:269`). **Keep Hot Paths on `Ctrl+1..0`**; put tab-jump on a non-conflicting chord (e.g. `Ctrl+Alt+1..9`) or rely on `Ctrl+Tab` cycling. Call this out in the Keyboard preferences.

**Acceptance criteria**
- New tab opens a second folder in the same pane without disturbing the other pane; closing returns to the previous tab.
- Each tab independently remembers folder, view/sort, filter, and history.
- Tabs (per side, with active index) survive restart.
- "Reopen closed tab" restores the last closed tab's folder.
- Copy/move/compare still operate against the *active* tab of the focused pane.
- New Commands self-test `cmd_pane_tabs_*` (open/close/switch/reorder/duplicate/reopen/persist).

**Effort:** L (touches the pane model, persistence, and every "current path" assumption). **Risk:** the "one folder per pane" assumption is threaded widely; audit all `GetCurrentPath`/`SetFolderPath` call sites (including plugins via IHost) to route through the active tab.

---

### G4 — True Command Palette `[P1]`

**Current state**
- No dedicated palette. The closest thing is the **F1 "Display Shortcuts"** window (`cmd/app/showShortcuts`): it has a search box and *can* run a command on Enter (`ShortcutsWindow.cpp:757` → `DispatchShortcutCommand`), but (a) it lists **only keyboard-bound** commands (`ShortcutsWindow.cpp:1273-1294`), so unbound commands like Invert Selection, Batch Rename, Make File List are unreachable; (b) matching is **substring, not fuzzy**; (c) it is framed as a shortcut reference, not a launcher.
- The project already *wants* this: `Notes_Scratch.md` twice asks for "a first-class entry point (menu/command palette)".

**Proposed design**
- New `cmd/app/commandPalette` (suggested `Ctrl+Shift+P`, alias `Ctrl+P`) opening a centered, themed, modeless overlay.
- Source **every** command from `CommandRegistry` (not just bound ones), showing display name, category, and current shortcut. Include dynamic/parameterized families where meaningful (themes, plugin actions, hot paths, known folders).
- **Fuzzy** ranking (subsequence match + score), with recent/frequent boosting. Enter runs the highlighted command against the focused pane; arrows navigate; Esc closes.
- Reuse the row model and `DispatchShortcutCommand` plumbing already proven in `ShortcutsWindow`; the palette is a superset (all commands + fuzzy + launcher framing), so consider factoring a shared command-index used by both.

**Acceptance criteria**
- Palette lists commands that have no keyboard binding and runs them.
- Fuzzy query (`bris` → "Brief", `batren` → "Batch Rename") ranks the intended command first.
- Running a pane command targets the focused pane; running an app command behaves identically to the menu path.
- Recently-used commands rank higher on reopen.
- New Commands self-test `cmd_app_commandPalette_*` (lists unbound command, fuzzy rank, executes, active-pane routing).

**Effort:** M. **Risk:** low — mostly a superset of existing shortcut-window infra; keep the palette read-through-registry so it can't drift from the command set.

---

### G5 — Fuzzy Go to Folder / Go to File `[P1]`

**Current state**
- No fuzzy quick-jump anywhere (repo search for "fuzzy" finds only an audit note + a vendored JS lib). The three adjacent features are each narrower: address-bar autocomplete matches **one path segment at a time** (`UI_NavigationView.md:1199-1201`); Quick Search (`Shift+Space`) matches **within the current folder only**; the Find window is the heavyweight modeless search.

**Proposed design**
- `cmd/pane/goToFolder` (suggested `Ctrl+G`): type partial, space-separated fragments; fuzzy-match against folder paths and jump the focused pane there on Enter. Source candidates from: recent/history, Hot Paths + Bookmarks (G8), mounted drives/known folders, and — when available — the **search service index** (`RedSalamanderSearchService`) for directories.
- `cmd/pane/goToFile` (suggested `Ctrl+P` if not taken by the palette; otherwise `Ctrl+Shift+G`): fuzzy-match files under the current tree (or indexed roots), open/focus on Enter.
- Both are modeless overlays reusing the G4 palette shell (same list/scoring widget, different data source). When the index is unavailable, degrade gracefully to a bounded background walk of the current tree (reuse Find's enumeration + cancellation).

**Acceptance criteria**
- `Ctrl+G` + fragments jumps to the best-matching folder without typing a full path or stepping level-by-level.
- `Ctrl+P`/go-to-file focuses the matched file in its folder (opening the folder if needed).
- Works with the search index when present; falls back to a bounded live walk otherwise, and never blocks the UI thread (see memory: broker-read-hangs-UI concerns).
- New Commands self-test `cmd_pane_goto_*`.

**Effort:** M (M+ if wiring the search-service index path). **Risk:** index availability and UI-thread safety — reuse existing async search plumbing; keep the fallback bounded.

---

### G6 — File Checksums (SHA-256 / MD5 / CRC32) `[P1]`

**Current state**
- No per-file checksum surfaced. The Properties path delegates to the shell (`FolderView.FileOps.cpp:1114` `SHObjectProperties(...)`), Compare compares **byte-by-byte** (`CompareDirectoriesEngine.cpp:3086` `memcmp`), and existing SHA/BCrypt usage is CI/build/Monitor-internal only.
- Note: the themed Properties dialog (`Alt+Enter`, described in `docs/UserGuide.md`) is the natural home for a hash section; confirm whether it is a distinct surface from the shell property sheet and add there.

**Proposed design**
- `cmd/pane/computeHash`: for the focused file (or each selected file), compute MD5, SHA-1, SHA-256, and CRC32 using BCrypt (already used in `RedSalamanderMonitor.cpp:1048`) for the crypto hashes; stream the file in chunks with progress + cancel for large files (off the UI thread).
- Present in a themed dialog with per-algorithm rows and a **Copy** button each, plus "compare against pasted value" (paste an expected hash → show match/mismatch). Also expose the values as a **Checksums** section in the themed Properties dialog.
- Optional: a Make-File-List field (`{sha256}`) so checksums can be exported for a selection (extends the existing Make File List macro set).

**Acceptance criteria**
- Selecting a file and running the command shows correct MD5/SHA-1/SHA-256/CRC32 (validate against `certutil`/known vectors).
- Large-file hashing streams with a cancellable progress indicator and never freezes the UI.
- Copy buttons place the exact lowercase hex on the clipboard; paste-to-compare reports match/mismatch.
- New Commands self-test `cmd_pane_computeHash_*` (known-vector correctness + cancel).

**Effort:** S–M. **Risk:** low; keep hashing async and chunked.

---

### G7 — First-Class HEIC/WebP/AVIF + Quick/Batch Image Convert `[P2]`

**Current state**
- ViewerImgRaw decodes non-RAW images via generic WIC (`ViewerImgRaw.Decode.cpp:721`), but its declared extension list (`ViewerImgRaw.Internal.h:86`) is `.bmp .dib .gif .ico .jpe/.jpeg/.jpg .png .tif/.tiff .wdp .jxr .hdp` — **no `.heic/.heif/.webp/.avif`**. HEIC appears only in a code comment; WebP/AVIF are unreferenced. They would open only if a system codec is installed *and* the user manually associates the extension.
- Conversion exists but only as **single-image** viewer export (`ViewerImgRaw.Export.cpp:629`, `Ctrl+S`) to PNG/JPEG/TIFF/BMP/GIF/JXR. No folder-view "convert" and no batch.

**Proposed design**
1. **Register modern formats.** Add `.heic .heif .webp .avif` to `IsWicImageExtension` and the default viewer associations, so they open out of the box when the platform/WIC codec is present. Detect missing codec and show a friendly, actionable message (link to install the HEIF/AV1/WebP Image Extensions), mirroring the VLC "install VLC" pattern.
2. **Batch convert command.** `cmd/pane/convertImage`: folder-view context-menu + command entry that converts the selected images to a chosen target (JPG/PNG/WebP), with quality slider (for lossy), output-folder choice (in place vs. subfolder), name templating, and overwrite policy. Reuse the WIC encode path from `ViewerImgRaw.Export` factored into a headless batch converter; run in the background with progress/cancel like other file operations.

**Acceptance criteria**
- `.heic/.webp/.avif` open in the image viewer when the codec is installed; a clear install prompt appears when it isn't.
- Converting a multi-selection of HEIC → JPG produces valid JPGs for every item with the chosen quality, honoring the overwrite policy, with progress/cancel.
- New self-tests: ViewerImgRaw decode registration + a Commands/FileOps `convertImage` batch test (fixture images, output validated).

**Effort:** M. **Risk:** codec availability varies by machine — never assume; gate on WIC probe and message clearly. AVIF/WebP encode support depends on installed WIC encoders; fall back to PNG/JPEG when a target encoder is absent.

---

### G8 — Bookmarks/Favorites Beyond 10 Slots + Places Sidebar `[P2]`

**Current state**
- The only bookmark-like store is **Hot Paths**, hard-capped at 10 (`Common/SettingsStore.h:360-371` `std::array<..., 10>`), surfaced only as a flat section in the drive/menu dropdown (`NavigationView.Menus.cpp:1856`). No general named-bookmark store, no favorites sidebar/tree.

**Proposed design**
1. **Unlimited named bookmarks** as a distinct store (keep Hot Paths as the fast 10 keyboard slots; bookmarks are the unlimited, organizable set):
   ```cpp
   struct Bookmark      { std::wstring id, displayName, path, iconGlyph; std::wstring group; };
   struct BookmarksSettings { std::vector<Bookmark> items; };
   ```
   Add/remove/rename/reorder in a Preferences → Bookmarks page and via a folder-view "Add to Bookmarks" command.
2. **Optional Places sidebar** (togglable, off by default to preserve the current look): a slim left panel listing **Quick access / Bookmarks**, **Hot Paths**, **Drives** (with capacity), **Known folders**, and **Connections** (saved profiles). Clicking navigates the focused pane; supports drop-to-bookmark. This also gives RedSalamander the "Locations + eject" affordance WhimFiles has in its sidebar.

**Commands / settings**
- `cmd/pane/addBookmark`, `cmd/app/bookmarks/manage`, `cmd/app/toggleSidebar`, `cmd/pane/goToBookmark/<id>`.
- New `settings.bookmarks`; sidebar visibility in `FolderLayoutSettings`.

**Acceptance criteria**
- Add arbitrary bookmarks (>10), grouped and reorderable, persisted across restart.
- Sidebar (when enabled) navigates the focused pane and can eject removable drives; hiding it restores the exact current layout.
- Hot Paths remain independent and keyboard-bound.
- New Commands self-test `cmd_app_bookmarks_*` and a sidebar visibility/persist test.

**Effort:** M (bookmarks) + M–L (sidebar is a new persistent surface). Ship bookmarks first; sidebar second. **Risk:** sidebar interacts with split-ratio/layout persistence — reuse `FolderLayoutSettings`.

---

### G9 — Hover Preview `[P2]`

**Current state**
- RedSalamander has a **docked** Preview Pane (`Alt+6`) and Thumbnails, but no hover-to-peek. WhimFiles shows image/PDF previews on hover without opening anything.

**Proposed design**
- Optional (Preferences-gated, with a hover delay) floating preview popup: on hover over an image/PDF/text row, after the delay, show a small themed preview near the cursor using the existing thumbnail/preview infrastructure and viewer plugins. Dismiss on move-away/scroll/keypress. Never block enumeration or the UI thread; cancel in-flight preview on selection change.

**Acceptance criteria**
- Hovering an image row shows a preview after the configured delay; moving away hides it.
- Setting off → no hover preview, no overhead.
- Works for at least images and PDFs; degrades silently for unsupported types.
- New self-test asserting hover-preview show/hide + setting gating (host-alert/observability stub pattern used by viewer tests — see memory `plan-015` observability trick).

**Effort:** M. **Risk:** perf/flicker on fast scroll — debounce and cancel aggressively.

---

### G10 — Always-Calculate Folder Sizes Column `[P2]`

**Current state**
- Folder sizes are on-demand: `Space` calculates one folder and advances; `Alt+F10` opens the ViewerSpace treemap; there's no mode that auto-fills a Size column for every folder in view.

**Proposed design**
- `FolderViewSettings` gains `alwaysCalculateFolderSizes` (per pane). When on, folder rows show a live-computed recursive size in the Size column, computed on a bounded background pool (reuse the icon/thumbnail thread-pool pattern), populated progressively, and cancelled/reset on navigate. Add a per-pane toggle in the sort/display popup.

**Acceptance criteria**
- With the mode on, folder Size cells fill in progressively without freezing scrolling; navigating away cancels pending work.
- Mode is per-pane and persists.
- Off by default (perf-safe); on huge trees the computation is bounded and cancellable.
- New self-test for the setting + progressive-fill behavior.

**Effort:** M. **Risk:** perf on deep trees / network paths — bound concurrency, skip/annotate slow remote FS, cancel on navigate (heed the project's perf-review culture).

---

### G11 — "Open in App" Terminal Helper `[P3]`

**Current state**
- RedSalamander does the *reverse* well ("Command Shell" / "Bring Current Directory to Command Line"), but has no documented shell helper to open a path *in RedSalamander* from the terminal, the way WhimFiles ships `wf` for zsh/bash.

**Proposed design**
- Confirm/ensure `RedSalamander.exe <path>` opens that folder in a pane (add if missing; support opening in the focused pane / a new tab once G3 lands).
- Ship a tiny `redsal`/`rs` PowerShell function + a `cmd`/bash shim that opens the current directory (or a given path) in RedSalamander, documented in `docs/`. Optional "Open in RedSalamander" shell context-menu entry via the installer.

**Acceptance criteria**
- `RedSalamander.exe "C:\some\path"` opens that folder (and focuses it).
- Documented one-liner shell helper opens the cwd in the app.
- Documented in `docs/UserGuide.md`.

**Effort:** S. **Risk:** minimal.

---

## 5. Non-Gaps (RedSalamander already matches or exceeds WhimFiles)

To keep scope honest, these WhimFiles features are **already covered** and need no work:

- **Session persistence** — RedSalamander already restores pane folders, view mode, sort, extension/nav/filter/status-bar visibility, split ratio, active pane, zoom, and history on restart (`RedSalamander.cpp:2821-2852` capture, `:9174-9345` restore). (Extend to tabs once G3 lands.)
- **Icon/thumbnail preview in lists** — Thumbnails mode (`Alt+5`) with background thumbnail loading + shell/WIC fallback.
- **Batch rename** — RedSalamander's is deeper (macros, search/replace, independent name/extension case, manual mode, live validation, modeless).
- **In-folder incremental search** — Quick Search (`Shift+Space`).
- **ZIP create** — Pack/Unpack (`Alt+F5`/`Alt+F6`), plus 7z virtual-FS browsing.
- **Conflict handling, background ops, speed limits, parallel mode** — richer than WhimFiles.
- **Docked preview + real viewers** — nine viewer plugins vs. Quick Look.

And RedSalamander's remote/cloud file systems, Compare+sync, plugin architecture, theming, and ETW monitor have **no WhimFiles equivalent** — these are moat features to protect, not gaps.

---

## 6. Recommended Roadmap

Sequence by value-per-effort and dependency:

**Wave 1 — daily-driver wins (do first)**
1. **G6 Checksums** (S–M, low risk) — quick credibility win, self-contained.
2. **G4 Command palette** (M, low risk) — unlocks discoverability; shared shell reused by G5.
3. **G1 Multi-dimensional filtering** (L) — the flagship; highest user-visible impact.

**Wave 2 — safety + navigation**
4. **G2 Undo** (M–L) — data-safety insurance; closes the `plans/013` spike.
5. **G5 Fuzzy go-to** (M) — builds on G4's overlay.
6. **G3 Tabs** (L) — biggest structural change; schedule when the pane-model refactor can be absorbed.

**Wave 3 — polish**
7. **G8 Bookmarks** (M) then Places sidebar (M–L).
8. **G7 HEIC/WebP/AVIF + batch convert** (M).
9. **G9 Hover preview** (M), **G10 auto folder sizes** (M), **G11 terminal helper** (S).

Each item ships behind the existing test discipline: RED→GREEN Commands/FileOps self-tests, schema tests for new settings, and docs/spec closeout (update `docs/UserGuide.md`, `Specs/UI/UI_FolderView.md`, `Specs/UI/UI_CommandMenuKeyboard.md`, and `Specs/Testing/Testing_TestCoverage.md`). Register new FileOps cases in `kFileOpsFamilyDefinitions` so Full runs don't silently skip them.

---

## 7. Open Questions

1. **Tabs vs. shortcut budget** — confirm the tab-jump chord given `Ctrl+1..0` is Hot Paths. Proposal: keep Hot Paths, put tab-jump on `Ctrl+Alt+1..9` + `Ctrl+Tab` cycling.
2. **`Ctrl+P` ownership** — command palette (G4) vs. go-to-file (G5). Proposal: `Ctrl+Shift+P` palette, `Ctrl+P` go-to-file (VS Code-consistent), `Ctrl+G` go-to-folder.
3. **Filter persistence granularity** — should saved multi-dimension filters be per-pane, per-tab (once tabs exist), and/or named/reusable "saved filters"? Recommend per-tab live state **plus** named saved filters in `settings.fileTypeGroups`-adjacent storage.
4. **Properties surface for checksums** — confirm the themed `Alt+Enter` Properties dialog is the right host vs. a standalone hash dialog (or both).
5. **Undo scope** — session-only (recommended) vs. persisted across restart (risky; source/destination state may have changed). Recommend session-only with a clear "cannot undo after restart" contract.

---

---

# Part II — Total Commander & File Pilot

Part I compared RedSalamander to WhimFiles, a *macOS* app. Total Commander and File Pilot are **Windows** dual-pane managers, so they are RedSalamander's most direct competitors and the fairest apples-to-apples benchmarks. They also come from opposite eras: Total Commander (1993→) is the mature, feature-maximal genre standard with a 30-year plugin ecosystem; File Pilot (2024→, public beta) is a from-scratch, performance-obsessed "next-gen" explorer. Between them they bracket the design space.

**Key finding:** both independently confirm the Part I gaps that matter most — **tabs (G3), command palette (G4), fuzzy navigation (G5), and rich filtering/flatten (G1)** — which promotes those from "nice ideas from a Mac app" to "table stakes for a modern Windows file manager." They then add seven power-user gaps (G12–G18) that are classic Total Commander parity items.

---

## 8. Total Commander

### 8.1 Positioning

| Dimension | RedSalamander | Total Commander |
|-----------|---------------|-----------------|
| Platform | Windows 11, native C++23 / Direct2D | Windows (Win32), 32/64-bit, also Android |
| Maturity | Young, in active development | ~30 years, extremely mature |
| Extensibility | Plugin API (FS + viewers), young | Huge ecosystem: packer / filesystem (WFX) / lister (WLX) / content (WDX) plugins, thousands available |
| Rendering | GPU Direct2D, themeable | Classic Win32 controls |
| Philosophy | Modern themeable successor to Open Salamander | Feature-maximal Swiss-army-knife |
| Pricing | Free / open | Shareware (~$/license) |

Total Commander is the feature ceiling of the genre. RedSalamander already matches or exceeds it on **rendering/theming, cloud/object-storage file systems (S3/OneDrive/SharePoint/Google Drive natively vs. TC-via-plugins), and modern viewers**, but TC has a long tail of power-user utilities RedSalamander has never built.

### 8.2 Comparison (distinguishing rows)

Rows already covered by G1–G11 are marked; new gaps are **G12–G18**. ✅ full · 🟡 partial · ❌ absent.

| Feature | RedSalamander | Total Commander | Gap |
|---|:--:|:--:|:--:|
| Folder tabs (+ restore on startup) | ❌ | ✅ | G3 |
| Live filter (type/size/date) / quick filter | 🟡 name mask | ✅ Ctrl+S quick filter | G1 |
| Branch view — flatten all files in subdirs | 🟡 via Find | ✅ Ctrl+B | G1 (+§10 note) |
| Directory hotlist quick-menu | 🟡 10 Hot Paths | ✅ Ctrl+D hotlist | G8 |
| File-op undo | ❌ | 🟡 limited | G2 |
| **Configurable button bar / toolbar** | ❌ | ✅ "Configurable button bar and Start menu" | **G12** |
| **Horizontal (top/bottom) panel arrangement** | ❌ side-by-side only | ✅ vertical + horizontal | **G13** |
| **Directory tree pane / dual tree** | ❌ | ✅ | **G14** |
| **Always-visible command line** | 🟡 transient popup | ✅ command line at bottom | **G15** |
| **Color files by type/attribute** | ❌ | ✅ color config by wildcard/attr | **G16** |
| **Split / combine big files** | ❌ | ✅ "Split/Combine big files" | **G17** |
| **Encode/decode (UUE/XXE/MIME/Base64)** | ❌ | ✅ | **G18** |
| **Create/verify checksum files (SFV/MD5/SHA)** | ❌ | ✅ | **G18** (+G6) |
| Compare files by content (editor) | ✅ ViewerText diff | ✅ built-in compare editor | — |
| Synchronize directories | ✅ Compare + sync | ✅ | — |
| Multi-rename tool | ✅ Batch Rename (regex-class) | ✅ (regex) | — |
| Thumbnails | ✅ | ✅ | — |
| FTP / network | ✅ FTP/SFTP/SCP + more | ✅ FTP (SFTP via plugin) | RS ≥ |
| Cloud / object storage | ✅ native S3/OneDrive/SharePoint/GDrive | 🟡 via plugins | RS wins |
| Content-plugin custom columns | 🟡 fixed columns | ✅ WDX custom columns | §10 minor |
| Rich internal viewers (hex/img/RAW/SQLite/PE/media/web) | ✅ 9 plugins | 🟡 Lister + plugins | RS ≥ |
| Themeable GPU UI | ✅ | ❌ classic Win32 | RS wins |

### 8.3 Total Commander takeaways

- **Shared gaps (already specced):** tabs (G3), filtering/flatten (G1), undo (G2), hotlist→bookmarks (G8).
- **New gaps (specced in §10):** button bar (G12), pane layout/orientation (G13), directory tree (G14), persistent command line (G15), file-type colors (G16), split/combine (G17), encode + checksum files (G18).
- **RedSalamander already ahead:** native cloud/object-storage FS, modern viewers, themeable GPU rendering, Compare+sync, Batch Rename depth.
- **Deliberately not chasing:** TC's full plugin-ABI compatibility (packer/WFX/WLX/WDX). RedSalamander has its own plugin API; re-implementing TC's ABI is out of scope. The one worth cherry-picking is **content-plugin-style custom columns** (derive extra list columns from provider metadata) — noted as a smaller item in §10.

---

## 9. File Pilot

### 9.1 Positioning

| Dimension | RedSalamander | File Pilot |
|-----------|---------------|-----------|
| Platform | Windows 11 (x64 + ARM64) | Windows 7+, x86-64 only |
| Footprint | Large solution + plugins | ~2 MB single exe, handmade custom engine |
| Status | In development | Public beta (~v0.7) |
| Pricing | Free / open | Free in beta; perpetual $50 (1yr updates) / $250 lifetime |
| Rendering | Direct2D | Fully custom-rendered (no OS controls) |
| Philosophy | Broad power tool | **Raw speed** first, keyboard-driven, minimal |
| Extensibility | Plugin API + 9 viewers | None advertised (no plugins/scripting/hex/media viewers) |

File Pilot's pitch is "faster than your thoughts": instantaneous tab open, millisecond search, buttery scrolling, seamless thumbnail scaling — all from a hand-written engine. It is *narrower* than RedSalamander (no remote/cloud FS, no rich viewers, no plugins, no compare/sync) but its interaction speed and modern UX are exemplary.

### 9.2 Comparison (distinguishing rows)

| Feature | RedSalamander | File Pilot | Gap |
|---|:--:|:--:|:--:|
| Folder tabs (+ persistence) | ❌ | ✅ | G3 |
| **Many panes, each with its own tabs** | ❌ two panes | ✅ "virtually no limit" to panes | G3 + **G13** |
| **Saveable / switchable layouts** | 🟡 session restore only | ✅ save & switch layouts | **G13** |
| Command palette (global action search, aliases, key sequences) | 🟡 F1 substring, bound-only | ✅ | G4 |
| GoTo / fuzzy quick-nav (recent + autocomplete) | ❌ | ✅ | G5 |
| Fuzzy search | ❌ | ✅ | G5 |
| Flattened hierarchy view across whole drives | 🟡 via Find | ✅ millisecond flatten | G1 (+§10 note) |
| Extension filtering | 🟡 name mask | ✅ | G1 |
| Inspector (peek text/images/folders) | ✅ Preview Pane + viewers | ✅ | RS ≥ (hover: G9) |
| Batch rename | ✅ deep | 🟡 basic (unique id, date) | RS wins |
| Customization: hotkeys / color themes / font & spacing / animation toggle | 🟡 themes + keys | ✅ | partial |
| Raw interaction speed as an explicit goal | 🟡 strong, not a headline | ✅ headline | §10 note |
| Remote / cloud FS, compare/sync, plugins, rich viewers | ✅ | ❌ | RS wins |

### 9.3 File Pilot takeaways

- **Validates the modern-UX core:** File Pilot independently ships tabs, a command palette, fuzzy GoTo, and instant flattened filtering — the exact Part I gaps (G3/G4/G5/G1). Two unrelated modern managers converging on the same four features is a strong signal to prioritize them.
- **New emphasis:** **multiple panes + saveable/switchable layouts** (feeds G13), and **performance as a stated product identity**. RedSalamander already invests heavily in perf (see the FolderView frame-performance and enumeration plans in `Specs/Plans/Done/`), so the gap is less raw speed than *making instant interaction a visible, defended product value* — e.g. instant tab open (G3), sub-frame filter updates (G1), and never blocking the UI thread on search (G5).
- **RedSalamander already ahead:** breadth (cloud/remote, viewers, compare, plugins, theming) that File Pilot explicitly does not attempt.

---

## 10. New Gap Specifications (G12–G18)

Same structure as G1–G11. All seven were **verified absent** in the current codebase (evidence inline). These are predominantly Total Commander parity utilities; most are low-risk and reuse existing infrastructure (command dispatch, the external-launch macro engine, the hashing core from G6, the color-picker from Themes, the G1 predicate engine).

---

### G12 — Customizable Button Bar / Toolbar `[P2]`

**Current state:** No user toolbar. `CommandRegistry.cpp:60-61` has only `cmd/app/toggleFunctionBar` / `cmd/app/toggleMenuBar`; `FunctionBar.h:23` is the fixed F1–F12 strip; no `ButtonBar` type exists. Nearest analog is the User Menu (`cmd/pane/userMenu`), a menu, not a toolbar.

**Proposed design:** An optional, togglable toolbar row whose buttons invoke **any** `CommandRegistry` command **or** launch an external program (reuse the `FileActionLauncher` macro engine already built for viewers/editors/User Menu — path/selection/opposite-pane macros). Each button: a Segoe Fluent glyph (already the menu-icon system), tooltip, and target (command id / program+args / submenu). Configure in a new **Preferences → Toolbar** page (add/remove/reorder/separator/import-export). Persist `settings.toolbar`.

**Commands:** `cmd/app/toggleToolbar`, `cmd/app/toolbar/manage`, dynamic `cmd/app/toolbar/button/<id>`.

**Acceptance:** add a button that runs a built-in command and one that launches a program with `{fullPath}` macros; both work; bar toggles; config persists across restart; new Commands self-test `cmd_app_toolbar_*`.

**Effort:** M. **Risk:** low (reuses dispatch + launcher + glyphs).

---

### G13 — Flexible Pane Layout: Orientation, Layouts, (N panes) `[P2]`

**Current state:** Side-by-side vertical split only (`FolderWindow.Layout.cpp:347-359` computes `leftWidth`/`_rightPaneRect` on the X axis; splitter arrows are left/right). Exactly two panes (`Pane{Left,Right}`). `FolderLayoutSettings` stores a single `splitRatio`. Session restore exists but there is no *named, switchable* layout concept. (TC has vertical **and** horizontal arrangement; File Pilot has many panes + saveable/switchable layouts.)

**Proposed design (staged):**
- **Stage A — orientation toggle (P2):** allow the two panes stacked top/bottom. Add `FolderLayoutSettings.orientation {SideBySide, Stacked}`; generalize `CalculateLayout` to split along the chosen axis; splitter + drag adapt. `cmd/app/togglePaneOrientation`. Small, high-visibility.
- **Stage B — saveable layouts (P2/P3):** save the whole window state (orientation, split ratio, per-pane folder/view/sort/filter, and tabs once G3 lands) as a **named layout**; switch instantly. Builds directly on the existing capture/restore path (`RedSalamander.cpp:2821-2852` / `:9174-9345`). `cmd/app/layouts/save`, `cmd/app/layouts/switch/<id>`.
- **Stretch — N panes (deferred):** File Pilot's "unlimited panes" is a deep model change (the binary Left/Right assumption is threaded through the engine and plugins). Recommend deferring; tabs (G3) already absorb most of the "juggle many folders" need.

**Acceptance:** Stage A toggles stacked/side-by-side and back with a working splitter in both axes, persisted. Stage B saves/switches ≥2 named layouts restoring all pane state. Self-tests `cmd_app_paneOrientation_*`, `cmd_app_layouts_*`.

**Effort:** Stage A M, Stage B L. **Risk:** layout-rect math has many consumers — audit all `CalculateLayout` rect users; do Stage A first.

---

### G14 — Directory Tree Pane `[P2]`

**Current state:** None (no `TreeView` in `FolderView`/`FolderWindow`; "tree" in the code is the Preferences category tree). TC has a directory tree (and dual-tree).

**Proposed design:** **Merge with the G8 Places sidebar into one left dock** rather than shipping two competing left panels. The dock has two modes/sections: **Places** (Bookmarks, Hot Paths, Drives, Known folders, Connections — from G8) and **Tree** (a lazy-loaded, virtualized folder hierarchy). Selecting a tree node navigates the focused pane; expanding lazy-enumerates via the active file system (works for local and, where cheap, virtual FS). Theme-aware, togglable, off by default.

**Commands:** fold into `cmd/app/toggleSidebar` with a tree section, plus `cmd/app/sidebar/mode/<places|tree>`.

**Acceptance:** tree shows/hides; expanding lazy-loads children without freezing; selecting navigates the focused pane; visibility + mode persist. Self-test `cmd_app_sidebar_tree_*`.

**Effort:** M–L. **Risk:** perf on huge/remote trees — lazy + virtualized + cancellable; **coordinate with G8 so there is a single left dock, not two.**

---

### G15 — Persistent Command-Line Bar `[P2]`

**Current state:** PARTIAL. A real command-line control exists (`FolderWindow.FileSystem.Navigation.Part.cpp:274` `CreateCommandLineControls`; runs via `LaunchCommandLine`), but it is **transient** — created hidden, summoned only by "Bring … to Command Line", auto-hides after Enter/Esc (`FolderWindow.Layout.cpp:335`, `:562-563`). TC keeps a command line always visible at the pane bottom.

**Proposed design:** Add an always-visible mode. `cmd/app/toggleCommandLine` keeps the bar docked at the focused pane's bottom, persists visibility in `FolderLayoutSettings`, runs in the focused pane's directory, and adds up/down **history recall**. Reuse the existing create/execute/launch plumbing verbatim; only add the persistent-visible state + history buffer.

**Acceptance:** toggle keeps the bar visible across navigation and restart; Enter runs in the focused pane's dir without hiding; ↑/↓ recall previous commands. Self-test `cmd_app_commandLine_persistent_*`.

**Effort:** S–M (plumbing already exists). **Risk:** low.

---

### G16 — File-Type Color Coding `[P2]`

**Current state:** None. The per-item text brush is chosen only from focus/selection state (`FolderView.Rendering.cpp:2480-2492`); the sole attribute-driven visual is hidden-file icon dimming (`:2410`). TC colors files by wildcard/attribute rules.

**Proposed design:** User-defined, ordered color rules `{matcher, textColor}` where `matcher` **reuses the G1 `FolderFilterCriteria`** (wildcard mask, type group, size/date/attribute predicate). First match wins; rules layer over the active theme. Evaluate once per item at enumeration time (not per frame) into a cached per-row brush index, then branch in the render brush selection before the default `_textBrush`. Configure in **Preferences → File Colors** using the existing Themes color-picker control. Persist `settings.fileColors`.

**Commands:** `cmd/app/fileColors/manage`.

**Acceptance:** rule `*.zip;*.7z → orange` colors those rows; ordering/first-match respected; selection highlight still wins on selected rows; rules persist. Self-test `cmd_app_fileColors_*`.

**Effort:** M. **Risk:** render hot path — precompile matchers (reuse `MaskSyntax`/G1), evaluate on enumeration, cache per row.

---

### G17 — File Split & Combine `[P2]`

**Current state:** None (verified — `.001`/split/chunk hits are floats/the pane splitter; pack/unpack are ZIP only). TC has "Split/Combine big files".

**Proposed design:** `cmd/pane/splitFile` splits the focused file into N-byte parts (`name.ext.001`, `.002`, … + a small manifest carrying original name, size, and per-part + whole-file SHA-256 from the G6 core), with a size dialog (presets: 1.44 MB, CD 700 MB, DVD 4.7 GB, FAT-safe, custom). `cmd/pane/combineFile` detects a `.001` set/manifest, recombines, and **verifies** the whole-file hash before reporting success. Runs on the background file-op engine (progress/cancel).

**Commands:** `cmd/pane/splitFile`, `cmd/pane/combineFile`.

**Acceptance:** split → combine yields a byte-identical original (hash-verified); a missing/corrupt part produces a clear, specific error, not a silent bad output. New FileOps self-test `FileOps_SplitCombine_*` (register in `kFileOpsFamilyDefinitions` or Full runs skip it).

**Effort:** M. **Risk:** low; integrity verification is mandatory, matching the project's data-safety culture.

---

### G18 — Encode/Decode + Checksum Files (SFV/MD5/SHA) `[P2]`

**Current state:** None (no base64/uue/sfv anywhere; `crc32` is internal to the ZIP path only). TC does UUE/XXE/MIME encode-decode **and** create/verify checksum files. This is the multi-file complement to G6's single-file checksum dialog.

**Proposed design:**
- **Encode/decode:** `cmd/pane/encodeFile` (Base64/MIME; UUE optional) and `cmd/pane/decodeFile` (auto-detect), background for large inputs.
- **Checksum files:** `cmd/pane/createChecksumFile` writes `.sfv` (CRC32), `.md5`, or `.sha256` for the current selection using the **shared hashing core from G6**; `cmd/pane/verifyChecksumFile` reads such a file and reports per-entry OK / mismatch / missing with a results list (reuse the File Operations issues surface).

**Commands:** `cmd/pane/encodeFile`, `cmd/pane/decodeFile`, `cmd/pane/createChecksumFile`, `cmd/pane/verifyChecksumFile`.

**Acceptance:** create `.sha256` for a multi-selection; tamper one file; verify reports exactly that file as mismatched and any absent file as missing; Base64 encode→decode round-trips byte-identically. Self-test `cmd_pane_checksumFiles_*` and `cmd_pane_encodeDecode_*`.

**Effort:** M. **Risk:** low; shares the G6 hashing core.

---

### §10 minor items (fold-ins, not standalone gaps)

- **Branch view / whole-drive flatten** (TC Ctrl+B; File Pilot flattened hierarchy): G1's recursive mode covers the core. Add an explicit `cmd/pane/branchView` for a filter-less flatten of the current tree, and ensure the recursive/flatten path can target a whole drive with the same instant feel File Pilot demonstrates.
- **Directory hotlist** (TC Ctrl+D): fold into G8 — expose bookmarks as a quick `Ctrl+D` popup in addition to the sidebar.
- **Content-plugin custom columns** (TC WDX): RedSalamander already has a plugin API; a future enhancement is letting a file-system/metadata provider contribute extra list columns. Small, niche — track separately from this plan.
- **Performance as a defended product value** (File Pilot): not a discrete feature but a north-star for how G1/G3/G5 must feel — instant tab open, sub-frame filter updates, never blocking the UI thread on search. Hold new features to that bar; lean on the existing frame-performance work in `Specs/Plans/Done/`.

---

## 11. Consolidated Three-Way Roadmap (supersedes §6)

Merging G1–G18 across all three competitors. Sequence by value-per-effort, dependency, and how many competitors validate each gap (✚ = number of the three that ship it).

**Wave 1 — modern-UX core (validated by all/most competitors; do first)**
1. **G4 Command palette** (M, low risk) ✚2 (WhimFiles, File Pilot) — shared shell reused by G5; upgrades existing F1 window.
2. **G1 Multi-dimensional filtering + flatten** (L) ✚3 (all three) — flagship; add whole-drive flatten (branch view) for TC/FP parity.
3. **G6 Checksums** (S–M, low risk) — quick self-contained win; hashing core is shared by G17/G18.

**Wave 2 — navigation, safety, structure**
4. **G5 Fuzzy Go to Folder/File** (M) ✚2 — builds on G4's overlay.
5. **G2 Undo** (M–L) — data-safety insurance; closes `plans/013`.
6. **G3 Tabs** (L) ✚3 (all three) — biggest structural change; unlocks G13 Stage B.

**Wave 3 — power-user parity (mostly Total Commander; low-risk batch)**
7. **G8 Bookmarks** + **G14 Directory tree**, delivered as **one left dock** (M, then M–L).
8. **G13 Stage A** horizontal/stacked orientation (M); **Stage B** saveable layouts (L, after G3).
9. **G12 Button bar** (M) · **G15 Persistent command line** (S–M) · **G16 File-type colors** (M).
10. **G17 Split/combine** (M) · **G18 Encode + checksum files** (M, shares G6 core).

**Wave 4 — polish**
11. **G7 HEIC/WebP/AVIF + batch convert** (M) · **G9 Hover preview** (M) · **G10 auto folder sizes** (M) · **G11 terminal helper** (S).

Every item ships behind the existing test discipline: RED→GREEN Commands/FileOps self-tests, `SettingsSchemaTests` for new settings, and docs/spec closeout (`docs/UserGuide.md`, `Specs/UI/UI_FolderView.md`, `Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/Testing/Testing_TestCoverage.md`). Register new FileOps cases in `kFileOpsFamilyDefinitions` so Full runs don't silently skip them.

---

## 12. Updated Open Questions (extends §7)

6. **One left dock vs. two panels** — confirm G8 (Places) and G14 (Tree) ship as a single toggleable dock with Places/Tree modes (recommended), not two separate left panels.
7. **N panes** — accept File Pilot-style unlimited panes as a long-term stretch, or treat tabs (G3) as the answer to "many folders at once" and keep the two-pane model? Recommend the latter for now.
8. **Toolbar vs. User Menu** — should G12's button bar and the existing User Menu (`cmd/pane/userMenu`) share one action/macro model? Recommend yes (both reuse `FileActionLauncher`).
9. **TC plugin ABI** — confirmed out of scope; only content-plugin-style **custom columns** are worth revisiting later via RedSalamander's own plugin API.

---

---

# Part III — The Premium Tier & The Broader Field

Parts I–II covered a Mac filtering app, the genre benchmark, and a speed-focused newcomer. Part III adds the **premium power tier** — Directory Opus and XYplorer — which sit *above* Total Commander on features and expose a deeper class of gaps around **file organization** (tags, ratings, labels), **automation** (scripting), and **rich metadata** (configurable columns). It then scans the rest of the field so coverage is complete.

**Key finding:** Directory Opus and XYplorer both ship six capabilities RedSalamander entirely lacks (all verified absent in code): per-file **tags/labels/ratings/comments (G19)**, **user scripting (G20)**, **duplicate finder (G21)**, a **true configurable multi-column grid with rich metadata (G22)**, **collections/saved searches (G23)**, and a **long tail of power tools (G24)**. G19 and G22 are the most strategically important — they move RedSalamander from "a fast browser of the filesystem" toward "a tool that helps you *organize* files," which is the premium tier's whole value proposition.

---

## 13. Broader Competitive Landscape

The full set of managers surveyed, and what each uniquely contributes to this analysis:

| Manager | Platform | Model | What it adds to the gap picture |
|---|---|---|---|
| **Directory Opus** | Windows | Premium all-in-one | Tags/ratings/labels, scripting, dup finder, metadata columns, dual trees, collections → **§14 (full)** |
| **XYplorer** | Windows | Portable scriptable power | Tags/color filters, scripting + custom events, dup finder, custom columns, catalog, branch view, sync-browse → **§15 (full)** |
| **Files** | Windows (WinUI 3) | Modern open-source | Colored tags, column/grid/adaptive layouts, cloud + git status, tabs → **§16** |
| **fman** | Win/Mac/Linux | Minimalist keyboard-first | Command palette + fuzzy + Python plugin/package system → **§16** (reinforces G4/G5/G20) |
| Double Commander | Win/Mac/Linux | Free open-source TC clone | Reinforces TC gaps (tabs, tools, custom columns, multi-rename) |
| FreeCommander XE | Windows | Free dual-pane | Reinforces tabs, tree, folder-size, basic power tools |
| Multi Commander | Windows | Free, scriptable, multi-pane | Reinforces scripting (G20) + multi-pane (G13) |
| Q-Dir | Windows | Quad-pane | Reinforces flexible/multi-pane layout (G13) |
| Far Manager | Windows | Text-mode, huge plugin ecosystem | Reinforces plugins + macros/scripting (G20) |
| One Commander | Windows | Modern, Miller columns | Reinforces Miller-column view + rich columns (G22), modern UI |
| Spacedrive | Cross-platform | Library/VDBMS, tags across devices | Reinforces tags/organization (G19) as a **library** — a long-term direction |
| Marta / Nimble Commander | macOS | TC-style dual-pane | Cross-platform validation of the dual-pane + tools model |
| Krusader / Dolphin | Linux | KDE dual-pane / single | Reinforces tags, split views, service-menu scripting |
| Yazi / ranger / nnn / lf | Terminal | Async, keyboard-first | Reinforces async-everything + instant preview (perf north-star) |
| muCommander | Java, cross-platform | Lightweight dual-pane | Minor; broad protocol support |

**Observation:** nothing in the wider field introduces a *seventh* new gap class beyond G19–G24 — they cluster around the same themes (tabs, palette, filtering, tags, scripting, columns, multi-pane), which increases confidence that G1–G24 is a **complete** map of what RedSalamander is missing relative to the market.

---

## 14. Directory Opus

### 14.1 Positioning

| Dimension | RedSalamander | Directory Opus |
|-----------|---------------|----------------|
| Platform | Windows 11, native C++23 / Direct2D | Windows, native 64-bit, multi-threaded |
| Maturity / market | Young, free/open | Mature commercial flagship (paid) |
| Identity | Themeable modern successor to Open Salamander | The maximal "do everything" file manager + Explorer replacement |
| Extensibility | Plugin API + 9 viewers | Full scripting interface (JScript/VBScript) + custom columns |

Directory Opus is the feature ceiling of the entire genre — it does everything Total Commander does *plus* file organization (tags/ratings/labels), a full scripting engine, a duplicate finder, a metadata editor, rich configurable columns, dual trees, and collections. It is the single best "north star" for where a maximal RedSalamander could go. RedSalamander already matches it on **native cloud/object-storage FS, GPU theming, and modern viewers**, and undercuts it on price, but trails on the organization/automation layer.

### 14.2 Comparison (distinguishing rows)

| Feature | RedSalamander | Directory Opus | Gap |
|---|:--:|:--:|:--:|
| Dual panes + **dual trees** + tabs | ❌ tabs / ❌ trees | ✅ all three | G3 / G14 |
| **Label, rate & tag files; descriptions** | ❌ | ✅ color labels, star ratings, tags, descriptions | **G19** |
| **Full scripting interface** (custom commands/dialogs) | ❌ | ✅ JScript/VBScript | **G20** |
| **Locate duplicates** (dedicated finder) | ❌ | ✅ | **G21** |
| **Metadata editor + rich/configurable columns** | ❌ fixed tile layout | ✅ metadata panel + columns | **G22** |
| Editable toolbars/menus + floating launchers | ❌ | ✅ "all toolbars and menus can be edited" | G12 |
| Integrate indexed search (Everything / Windows Search) | 🟡 own search service | ✅ | §17 note |
| Queued file copies | ✅ Wait/Parallel + speed limit | ✅ queued copies | — |
| Batch rename (macro/keyboard) | ✅ deep | ✅ | — |
| Synchronize / backup | ✅ Compare + sync | ✅ | — |
| Image viewer + marking + **slideshow** | 🟡 ImgRaw viewer | ✅ marking + slideshow | G7 note |
| 600+ configurable colors/fonts; file-type colors | 🟡 themes | ✅ + color-by-type | G16 |
| FTP + archives (Zip/7z/RAR) | ✅ FTP/SFTP/S3/cloud + 7z/ZIP | ✅ | RS ≥ (cloud) |
| Native cloud/object storage (S3/OneDrive/SharePoint) | ✅ | 🟡 add-ons | RS wins |
| Themeable GPU rendering | ✅ | 🟡 configurable but classic | RS wins |

### 14.3 Directory Opus takeaways
- **New gaps:** file organization (G19), scripting (G20), duplicate finder (G21), configurable metadata columns + metadata editor (G22), collections (G23).
- **Reinforces:** tabs (G3), dual tree (G14), editable toolbars (G12), color-by-type (G16).
- **RedSalamander already ahead:** native cloud/object-storage, GPU theming, price.
- **Strategic read:** DOpus's differentiator over TC is *organization + automation*. That is exactly the G19/G20/G22 cluster — the highest-value place for RedSalamander to grow beyond "a fast browser."

---

## 15. XYplorer

### 15.1 Positioning

| Dimension | RedSalamander | XYplorer |
|-----------|---------------|----------|
| Platform | Windows 11, native | Windows, fully **portable** (no install) |
| Identity | Themeable modern successor | Fast, scriptable, tagging-centric power tool |
| Signature | Cloud/remote FS + viewers + theming | Tagging + color filters + scripting + catalog |

XYplorer overlaps DOpus on the organization/automation axis but is portable and scripting-forward. It is the strongest reference for **tagging + color filters + a scripting language + custom columns + branch view + sync-browse** done in a lightweight package.

### 15.2 Comparison (distinguishing rows)

| Feature | RedSalamander | XYplorer | Gap |
|---|:--:|:--:|:--:|
| Tabs + Mini Tree + breadcrumb + drive bar | 🟡 breadcrumb only | ✅ all | G3 / G14 |
| **Labels, tags, comments (multi-user)** | ❌ | ✅ | **G19** |
| **Color filters** (color by date/size/tag/type) | ❌ | ✅ | G16 + **G19** |
| **Full scripting language + custom event actions + user commands** | ❌ | ✅ | **G20** |
| **Duplicate finder** (content hash / image similarity) | ❌ | ✅ | **G21** |
| **Custom columns** (user-defined, scriptable, incl. GPS/EXIF) | ❌ fixed | ✅ | **G22** |
| **Catalog** (file catalog independent of location) | ❌ | ✅ | **G23** |
| Live filter box (filter-as-you-type) | 🟡 name mask | ✅ | G1 |
| Branch view (flat view) | 🟡 via Find | ✅ | G1 |
| Search-results caching | ❌ | ✅ | G23 |
| Verified (hash) file copying | 🟡 partial verify | ✅ | G6 note |
| **Hash values (MD5/SHA) display** | ❌ | ✅ | G6 |
| Sync browse / sync scroll | ❌ | ✅ | **G24** |
| Secure delete / wipe | ❌ | ✅ | **G24** |
| Manual/custom & secondary sort; selected-to-top | 🟡 fixed sort keys | ✅ | G24 |
| Preview pane + quick audio/video/photo (zoom/histogram/RAW) | ✅ viewers + preview | ✅ | — |
| Batch rename with preview | ✅ | ✅ | — |
| Portable install | 🟡 MSIX/MSI | ✅ fully portable | minor |

### 15.3 XYplorer takeaways
- **New gaps:** same core cluster as DOpus — tags/labels/comments (G19), scripting + custom events (G20), duplicate finder (G21), custom columns (G22), catalog/collections (G23) — plus the **G24 long tail** (sync-browse/scroll, secure wipe, manual sort, hash display).
- **Reinforces:** tabs (G3), tree (G14), color-by-type/filters (G16), branch view + live filter (G1), checksums (G6).
- Two independent premium products converging on the identical G19–G23 set strongly validates prioritizing them.

---

## 16. Files (files.community), fman & the field

**Files** — modern open-source WinUI 3 / Fluent explorer replacement. Adds: **colored file tags (G19)**, multiple **column/grid/adaptive layouts (G22)**, tabs + dual pane (G3), cloud-drive status + **git integration**, QuickLook preview, archive handling. Represents the *modern-Windows-native* direction; free/open. Reinforces G19 and G22 with a Fluent-design reference RedSalamander (also Fluent-leaning) can learn from. *(Site returned 403 to automated fetch; characterized from the public feature set + reviews.)*

**fman** — cross-platform (Win/Mac/Linux), minimalist, keyboard-first. Adds: **command palette + fuzzy GoTo (`Ctrl+P`) (G4/G5)**, smart folder suggestions, and a **Python plugin + package-manager system (G20-adjacent)**. Reinforces that a command palette, fuzzy navigation, and a *scriptable* plugin/package model are the modern baseline.

**The rest** (Double Commander, FreeCommander XE, Multi Commander, Q-Dir, Far Manager, One Commander, Spacedrive, Marta/Nimble, Krusader/Dolphin, Yazi/ranger/nnn/lf, muCommander) each reinforce gaps already captured — see the §13 table. Two are worth a forward-looking note:
- **One Commander** — a polished **Miller-columns** browsing mode (feeds the G22 view-modes discussion).
- **Spacedrive** — a **library/VDBMS** model that indexes files with tags *across devices*. A long-term vision for where G19 (tags) + the existing search service could converge; out of scope now, noted for direction.

---

## 17. New Gap Specifications (G19–G24)

Same structure as before. All verified absent (evidence inline from the codebase audit). Several reuse infrastructure already specced: the G6 hashing core (G21), the search-service SQLite store (G19/G23), the G1 predicate engine + G16 rendering (G19 display/filter), and the command palette (G20 custom commands).

---

### G19 — File Tags, Labels, Ratings & Comments (per-file organization) `[P1]`

**Why it matters:** This is the premium tier's core differentiator (DOpus, XYplorer, Files all ship it) and the biggest step from "browsing the filesystem" to "organizing files." Highest strategic value in Part III.

**Current state:** No per-file user metadata anywhere (verified). The item model carries only `displayName/isDirectory/sizeBytes/lastWriteTime/fileAttributes` (`FolderView.h:691`); the search index `entries` table is filesystem-only (`SqliteIndexStore.cpp:1286-1304`); no tag/rating/label/comment command in the 126-command registry; the only `label` field is a bookmark display name (`SettingsStore.h:363`).

**Proposed design**
- **Annotation store.** A SQLite table (reuse the search-service store infrastructure) keyed by a **stable file identity** — volume serial + file ID for local NTFS (survives rename/move on the same volume), path-based fallback for remote/virtual FS. Fields: tags (many-to-one, from a named+colored **tag catalog**), a 0–5 **rating**, a single **color label**, and a free-text **comment/description**.
- **Assign:** context-menu + commands + a section in the themed Properties dialog. Bulk-assign to a multi-selection.
- **Display:** as row tint/badge (reuse G16 render path) and as columns (rating, label, tags — requires the G22 grid for full column display; until then show in the details/metadata line and Properties).
- **Filter & sort:** extend the G1 `FolderFilterCriteria` with tag/rating/label dimensions ("show 4★+ / tagged #invoice"); add sort-by-rating.
- **Search:** extend the index so tags/comments are searchable in the Find window.

**Commands / settings:** `cmd/pane/tags/assign`, `cmd/pane/rate/<0-5>`, `cmd/pane/label/<id>`, `cmd/pane/setComment`, `cmd/app/tags/manage`; `settings.tagCatalog` + a tags DB.

**Acceptance:** tag/rate a file; it shows a badge and persists across restart; rename/move on the same volume keeps the tags; filter to `rating≥4` and to a tag; sort by rating; search finds a file by tag. Self-test `cmd_pane_tags_*` + a persistence/identity test.

**Effort:** L (new persistent store + identity tracking + several UI surfaces). **Risk:** identity stability across move/rename (use NTFS file-id; degrade gracefully on FAT/remote); orphan cleanup when files vanish. Reuses search-service SQLite infra.

---

### G20 — User Scripting / Automation `[P2]`

**Current state:** No end-user scripting/macro engine (verified). No script host (`IActiveScript`/Lua/Python) in the tree or `vcpkg.json`; "macro" means Batch-Rename tokens and Make-File-List templates; the plugin API is C++ COM. DOpus (JScript/VBScript), XYplorer (scripting + custom event actions + user commands), fman (Python) all have it.

**Proposed design**
- Embed a scripting host exposing a stable **automation API**: navigate panes, enumerate/select items, run file operations *through the existing engine* (so undo/G2, conflicts, and progress all apply), read/write tags (G19), and register **custom commands** that appear in the command palette (G4) and are bindable to keys (Keyboard prefs) and toolbar buttons (G12).
- **Engine choice:** recommend **JScript via Windows `IActiveScript`** first — zero new runtime dependency, native, and directly familiar to the large Directory Opus scripting community; consider Lua/QuickJS later for a sandboxed modern option. The IHost automation surface already sketched in `Notes_Scratch.md` (Navigate / FocusItem / ChangeSelection / GetSelection across panes, VFS-aware) is the natural API spine.
- **Custom event actions** (XYplorer-style): optional hooks (on-folder-change, on-startup) gated by a setting.

**Commands / settings:** `cmd/app/runScript`, `cmd/app/scripts/manage`; user scripts register `cmd/user/<id>`; a Preferences → Scripting page with an enable/consent toggle.

**Acceptance:** a script that selects `*.tmp` and deletes them (via the engine, so it's undoable) runs from the palette; a user-defined command binds to a key and appears in the palette; a script reads and sets a tag. Self-test `cmd_app_script_*`.

**Effort:** L. **Risk:** security — scripts can mutate files; gate behind explicit enable + consent, run through the audited engine, and never auto-run untrusted scripts. Schedule deliberately (large, cross-cutting).

---

### G21 — Duplicate File Finder `[P2]`

**Current state:** None (verified — `selectSameName`/`selectSameExtension` only match names within one folder). DOpus "Locate duplicates"; XYplorer duplicate finder by content hash / image similarity.

**Proposed design:** `cmd/app/findDuplicates` scans one or more roots, **buckets by size, then confirms by content hash** (reuse the G6 hashing core), and presents groups in the Find results window with safe group actions: keep-newest / keep-oldest / keep-one, with a hard **never-delete-the-last-copy** guard and normal delete confirmations. Background + cancel. Optional **image-similarity** (perceptual hash) mode as a later stretch.

**Commands:** `cmd/app/findDuplicates`.

**Acceptance:** a folder with known duplicates groups correctly by content (not just name); bulk delete never removes the last copy of a group without explicit confirm; a large scan is cancellable. Self-test in the FileOps/Commands family (register in `kFileOpsFamilyDefinitions`).

**Effort:** M. **Risk:** data-safety on bulk delete (route through the engine + guards; pairs naturally with G2 undo); perf (size-bucket before hashing so most files are never hashed). Shares the G6 core.

---

### G22 — True Multi-Column Grid + Rich, Configurable Metadata Columns `[P1/P2]`

**Why it matters:** This is both a UX gap *and* a structural limitation — and it unblocks G19's column display. Every premium/modern manager (DOpus, XYplorer, Files, One Commander) has a real column grid with a column chooser and metadata columns; RedSalamander's "Detailed" view is a **tile layout**, not a grid.

**Current state (verified):** `DisplayMode = {Brief, Detailed, ExtraDetailed, Thumbnails}` (`FolderView.h:348-354`). Detailed/ExtraDetailed render a fixed `label + details-line (time•size•attributes) + optional metadata-line`, not a resizable multi-column grid (`FolderViewInternal.h:419-441`). No column chooser (the `DxUi.Grid` control exists but is used only inside dialogs). Sort keys are limited to `Name/Extension/Time/Size/Attributes/None` (`FolderView.h:356-364`); the metadata callback exposes only name/size/date/attr — no EXIF, dimensions, duration, or audio tags.

**Proposed design (staged)**
- **Stage A — real column grid (P1/P2):** a Details view that is a genuine resizable + reorderable multi-column grid with a **column chooser** and click-to-sort headers. Evaluate generalizing the existing `Common/DxUi/DxUi.Grid` control (already used in dialogs) for the FolderView vs. evolving the tile layout. Built-in columns: Name, Ext, Size, Type, Modified, Created, Accessed, Attributes, Path. Persist column set/order/width per view in `FolderViewSettings`.
- **Stage B — rich metadata columns (P2):** a lazy per-row metadata provider computing optional columns on the background pool (like icons/thumbnails, cancel on navigate): image dimensions + EXIF (reuse WIC/libraw already in-tree), media duration/codec, audio/ID3 tags, PE version info, owner. Expose as optional columns *and* sort keys; allow plugins to contribute columns (the Total Commander "content column" idea from §10).

**Commands / settings:** `cmd/pane/columns/choose`; dynamic sort by any active column; `settings` per-view column layout.

**Acceptance:** user adds a "Dimensions" column; images show W×H and sort by it; adds/removes/reorders/resizes columns and the layout persists per view; metadata computes lazily without stalling scroll; folders with no metadata degrade cleanly. Self-tests for grid layout + metadata provider.

**Effort:** Stage A **L** (the biggest structural item — touches render/layout core), Stage B M–L. **Risk:** high surface area; do Stage A behind the existing frame-performance discipline (`Specs/Plans/Done/` FolderView perf plans). High value: it is the platform for G19 columns and closes the single biggest "feels less capable than DOpus/Files" impression.

---

### G23 — Collections / Saved Searches / Virtual Folders `[P2/P3]`

**Current state:** None (verified). Save/Restore Selection is a single, unnamed, **session-only** in-memory slot applied by name in one folder (`FolderWindow.h:1656-1665`); Make File List emits static text. DOpus collections; XYplorer catalog + search-result caching.

**Proposed design:** Persistent named **collections** — a virtual folder holding references to arbitrary files from anywhere; populate by adding a selection or by "save these search results." Implement as a **`collection:` virtual file-system provider** (RedSalamander already has the plugin-FS concept), so a collection is browsable, sortable, and operable like a folder, and appears in the Places sidebar (G8). Add **saved searches** — persisted Find queries re-runnable on demand. Store in settings/DB; handle missing items gracefully.

**Commands:** `cmd/pane/addToCollection`, `cmd/app/collections/manage`, `cmd/pane/saveSearch`.

**Acceptance:** add files from two different folders to a named collection, revisit it after restart, and operate on items; save a Find query and re-run it later. Self-test `cmd_app_collections_*`.

**Effort:** M–L. **Risk:** staleness of referenced items (show + prune missing). Lower priority than G19–G22.

---

### G24 — Power-User Long-Tail Bundle `[P2/P3]`

A batch of individually small XYplorer/DOpus utilities, grouped because each is low-risk and cheap once the surrounding systems exist:

- **Sync browse / sync scroll** — both panes navigate and/or scroll together (`cmd/app/toggleSyncBrowse`, `cmd/app/toggleSyncScroll`). Great for diffing parallel trees.
- **Secure delete / wipe** — overwrite before delete (`cmd/pane/wipeDelete`), with clear "irreversible" warning.
- **Manual / custom + secondary sort** — drag-to-reorder persisted per folder; multi-level sort; "selected to top" (extends the sort model).
- **EXIF/metadata rename tokens** — date-taken, dimensions in Batch Rename (extends its existing macro set; leans on G22's metadata provider).
- **Directory print / HTML report** — extend Make File List with a printable/HTML report output.
- **Drop stack** — a temporary holding area to gather files across folders before one copy/move (`cmd/app/dropStack`).

**Acceptance:** each ships with its own focused self-test. **Effort:** S each (M total). **Risk:** low; secure-delete must warn clearly and route through the engine.

---

## 18. Final Consolidated Roadmap (G1–G24) — supersedes §11

Sequenced by value-per-effort, dependency order, and competitor validation (✚ = independent competitors shipping it). Foundational items that unblock others are pulled earlier.

**Wave 1 — modern-UX core (validated across the field)**
1. **G4 Command palette** (M) ✚3 — reused by G5/G20.
2. **G1 Multi-dimensional filtering + flatten/branch view** (L) ✚5 — flagship interaction.
3. **G6 Checksums** (S–M) — self-contained; hashing core shared by G17/G18/G21.

**Wave 2 — structure & safety (unlock later waves)**
4. **G22 Stage A — real column grid + column chooser** (L) ✚4 — *foundational*: platform for rich columns and G19's tag/rating columns; biggest single UX credibility gain.
5. **G5 Fuzzy Go to Folder/File** (M) ✚2 — builds on G4.
6. **G2 Undo** (M–L) — data-safety; pairs with G21.
7. **G3 Tabs** (L) ✚6 — most-validated gap in the whole study; unlocks G13 Stage B.

**Wave 3 — organization (the premium-tier differentiator)**
8. **G19 Tags / labels / ratings / comments** (L) ✚4 — highest strategic value; display leans on G22, filter on G1, render on G16.
9. **G22 Stage B — rich metadata columns** (M–L) — completes the grid; feeds G19 + G24 rename tokens.
10. **G21 Duplicate finder** (M) ✚4 — shares G6 core; pairs with G2.
11. **G8 Bookmarks + G14 Directory tree** as one **left dock** (M → M–L).

**Wave 4 — power-user parity & automation**
12. **G13** orientation (M) + saveable layouts (L, after G3).
13. **G12 Button bar** (M) · **G15 Persistent command line** (S–M) · **G16 File-type colors** (M).
14. **G17 Split/combine** (M) · **G18 Encode + checksum files** (M).
15. **G20 User scripting** (L) — large, security-sensitive; do after the automation API (IHost surface) and G12/G4 exist so custom commands have a home.
16. **G23 Collections / saved searches** (M–L) · **G24 long-tail bundle** (opportunistic).

**Wave 5 — polish**
17. **G7 HEIC/WebP/AVIF + batch convert** (M) · **G9 Hover preview** (M) · **G10 auto folder sizes** (M) · **G11 terminal helper** (S).

Test discipline unchanged: RED→GREEN Commands/FileOps self-tests, `SettingsSchemaTests` for new settings, and docs/spec closeout. Register new FileOps cases in `kFileOpsFamilyDefinitions`.

**The three strategic bets, if forced to pick:** **G3 (tabs)** — the most-validated gap; **G1 (filtering)** — the highest-frequency interaction; **G19 + G22 (tags + real columns)** — the move up into the premium "organization" tier. Everything else is parity or polish around those.

---

## 19. Final Open Questions (extends §7 and §12)

10. **Tag storage** — SQLite sidecar (portable across FS, needs identity tracking) vs. NTFS alternate data streams (travels with the file but local-NTFS-only) vs. both. Recommend SQLite-primary with optional ADS export.
11. **Grid strategy (G22)** — generalize the existing `DxUi.Grid` for the FolderView, or evolve the current tile layout into columns? This is the pivotal architectural decision of the whole plan; prototype both before committing.
12. **Scripting engine (G20)** — JScript via `IActiveScript` (no new dep, DOpus-familiar) vs. embedded Lua/QuickJS (sandboxed, modern). Recommend JScript first.
13. **Collections as a VFS** — confirm implementing G23 as a `collection:` plugin file system (reuses the pane/engine) rather than a bespoke UI.
14. **How far up-market?** — G19/G20/G22 move RedSalamander into Directory Opus territory. Confirm that "organization + automation" is a desired product direction before committing Wave 3–4, or keep RedSalamander lean and stop at Wave 2.

---

*End of plan. This document is analysis + specification across the competitive field (WhimFiles, Total Commander, File Pilot, Directory Opus, XYplorer, Files, fman, and the broader landscape); no production code has been modified. Gaps G1–G24 are, on the evidence gathered, a complete map of RedSalamander's feature deltas versus the market.*
