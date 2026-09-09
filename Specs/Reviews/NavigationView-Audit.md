> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


I'll produce the final report directly from the verified findings. No code exploration is needed since findings are already verified with file:line evidence.

# NavigationView Audit — Final Report

## 1. Verdict

The address bar is structurally unsafe for the UI thread and sloppy about path semantics, but it is not a code-execution surface. The single most pervasive defect is **synchronous, uncancellable filesystem/shell I/O on the UI thread** — it recurs in at least five independent code paths (path-commit validation, per-keystroke plugin mount, drive-menu enumeration, menu-icon resolution, and the full-path-popup commit) and any one of them freezes the entire window for the full SMB/DNS timeout (tens of seconds) on a single offline UNC path or dead mapped drive. The second theme is **path-semantics correctness**: drive-relative input and unconditional `%VAR%` expansion silently navigate to the wrong target, and the full-path popup validates the pre-expansion literal so it rejects valid env-var paths. The third theme is **stale-async / lifetime hygiene**: autocomplete results and popups survive dismissal without invalidating the in-flight request, and a posted plugin-navigate replays against a possibly-swapped plugin. A fourth, lower theme is **menu-ID contract rot**: sibling IDs overflow their documented 600-699 range. Good news: no confirmed crash, use-after-free, or remote-code-execution finding survived verification — the env-var "injection" framing was explicitly refuted, and the popup-teardown UAF degraded to "menu-session corruption" guarded by null checks. The danger is real but bounded to hangs and wrong-navigation, not memory corruption.

---

## 2. Findings by Severity

### WRONG-TARGET (navigate to a location the user did not intend)

**Drive-relative input (`C:foo`) resolves against the process CWD, not the displayed folder**
`RedSalamander/NavigationView.Edit.cpp:2280` — wrong-target — **CONFIRMED**
`ValidatePath` gates the plugin-prefix branch on `colon >= 2`, so `C:foo` (colon at index 1) falls through to `GetAttributes("C:foo")` → `GetFullPathNameW`, which Windows resolves against the *per-drive current directory* of the process (typically the install/working dir). The raw typed text — not the canonicalized form — is then navigated via `RequestPathChange` (`NavigationView.cpp:325`), so it is re-resolved against the per-drive CWD at enumeration time. `parseWindowsPath` (`NavigationViewInternal.h:472`) interprets the same input as drive-*root* + filter, so the suggestion display and the actual navigation target diverge.
Scenario: typing `C:Windows` lands in `<app-cwd-on-C>\Windows`, not `C:\Windows`.
Fix: detect drive-relative input (alpha + `:` not followed by a separator) explicitly and either reject it with the invalid-path message or canonicalize it against the currently displayed folder/drive root *before* validating and navigating.

**Posted plugin navigate request is replayed with no re-validation of the active plugin**
`RedSalamander/NavigationView.cpp:709` — wrong-target — **PLAUSIBLE**
`NavigationMenuRequestNavigate` validates the cookie against `_fileSystemPlugin.get()` at *post* time (`:303`) and PostMessages only the raw, plugin-relative path text (`:316`). `OnNavigationMenuRequestPath` (`:709`) unconditionally calls `RequestPathChange(path(*text))` with no re-check. `SetFileSystem` (`:1298/1307`) — a synchronous UI-thread call — can run between the post and the dispatch, replacing `_fileSystemPlugin`/`_pluginShortId`/`_navigationMenu`. A bare `/some/folder` posted by archive plugin A is then resolved against plugin B's namespace (or, for the file plugin, becomes a drive-relative/garbage location). RAII-owned payload, so no UAF — purely mis-navigation within the app's own namespace.
Fix: capture the issuing plugin identity (cookie or `pluginShortId`) in the posted payload; in the handler, drop the request if `_fileSystemPlugin` is no longer the same instance, or re-qualify the path with the captured plugin prefix before navigating.

**Unconditional `%VAR%` expansion silently retargets navigation and makes literal `%TOKEN%` folders unreachable**
`RedSalamander/NavigationLocation.h:299` (also `NavigationView.Edit.cpp:861`) — wrong-target / incorrect-behavior — **CONFIRMED**
`NormalizeUserTypedLocationText` runs `ExpandEnvironmentStringsW` whenever a `%` is present, with no escape syntax and no "does the literal path exist?" check. `ValidatePath` then runs against the already-expanded text, so it can never fall back to the literal. A real folder named `%TEMP%` / `%USERPROFILE%` is silently rewritten to the expanded location and becomes unreachable from the address bar.
Note: the "host-controlled target / injection" framing is **REFUTED** — `ExpandEnvironmentStringsW` does pure string substitution from the process's own environment block; it launches nothing. Also the `100%foo → 100` mangling sub-claim is wrong: unpaired `%` is left verbatim. Only balanced `%DEFINED_VAR%` tokens are affected, so real-world reachability is low.
Fix: only expand when the literal (unexpanded) path does not resolve, or restrict expansion to a leading `%VAR%` segment, or provide an escape. At minimum, prefer the literal path when `GetFileAttributes` on the unexpanded text succeeds.

---

### DoS / UI-THREAD FREEZE

> **Same root cause, five reachable entry points.** The fix is one architectural change applied everywhere: never call blocking filesystem/shell APIs (`GetFileAttributesW`, `GetVolumeInformationW`, `GetDiskFreeSpaceExW`, `GetDriveTypeW`, `SHGetFileInfoW` without `SHGFI_USEFILEATTRIBUTES`, `LoadLibraryExW`+`Initialize`) synchronously on the UI thread. The codebase already has the worker-thread + `PostMessagePayload` pattern (edit-suggest worker, `_driveInfoCv` worker) — these paths simply don't use it.

**Address-bar / full-path-popup commit blocks on synchronous `GetFileAttributesW` for UNC / offline paths**
`RedSalamander/NavigationView.Edit.cpp:2320` (popup entry at `NavigationView.FullPathPopup.cpp:1185`) — dos-oom — **CONFIRMED**
This is the headline freeze, reachable from both the inline breadcrumb editor (`ExitEditMode`, `:863`) and the full-path popup editor (`ExitFullPathPopupEditMode`, `:1185`). For the file plugin, `ValidatePath` falls through to `_fileSystemIo->GetAttributes(...)` → `FileSystem::GetAttributes` (`Plugins/FileSystem/FileSystem.DirectoryOps.cpp:510`) → blocking `GetFileAttributesW`, with no timeout/cancellation/worker offload. `TryGetUncServerRoot` (`FileSystem.Path.cpp:111`) only short-circuits *bare* server roots (`\\server\`); a share/sub-path like `\\deadhost\share\dir` falls through to the blocking call. Suggestion enumeration is correctly async (`EditSuggestWorker`), proving the off-thread mechanism exists but is unused on commit.
Scenario: paste `\\10.0.0.99\share` and press Enter → UI frozen for the full SMB/DNS/connect timeout (tens of seconds); user can't even Esc (no messages pumped). Violates the CLAUDE.md no-UI-thread-blocking rule.
Fix: accept syntactically-valid absolute paths optimistically and let the async navigation/enumeration pipeline surface the error (mirroring the non-file-plugin branch at `:2303-2307`), or move the `GetAttributes` probe onto the edit-suggest worker with a bounded timeout and post the result back.

**Drive/menu dropdown enumerates all volumes synchronously on the UI thread**
`Plugins/FileSystem/FileSystem.Menu.cpp:264` (entry at `NavigationView.Interaction.cpp:228`) — dos-oom — **CONFIRMED**
A single click on the drive/menu button → `ShowMenuDropdown` (`NavigationView.Menus.cpp:879`) → `FileSystem::GetMenuItems`, which loops every bit of `GetLogicalDrives()` and for *each* drive calls `GetVolumeInformationW` (`:18`) and `GetDiskFreeSpaceExW` (`:25`) — no `DRIVE_REMOTE`/`DRIVE_CDROM` skip, no timeout, and the whole loop holds `_stateMutex`. The disk-info dropdown adds a synchronous `GetDriveTypeW` (`:436`). The authors already offload the equivalent `GetDriveInfo` to a worker (`NavigationView.Rendering.cpp:1108`), proving they know these block.
Scenario: a `Z:` mapped to a dead SMB host or a spun-down USB drive freezes the message pump for the full network timeout per drive *before the menu even appears*.
Fix: build the drive menu from a cached snapshot maintained by the existing background worker, or enumerate with a per-drive timeout and a "(not ready)" placeholder; at minimum skip volume/free-space queries for `DRIVE_REMOTE`/`DRIVE_CDROM`.

**`UpdateMenuIconBitmap` does live shell I/O (`SHGetFileInfoW`) on the UI thread for UNC/offline roots**
`RedSalamander/NavigationView.Rendering.cpp:519` — dos-oom — **CONFIRMED**
The menu-icon lookups pass `useFileAttributes=false` (`:435/490/519/526`), so `IconCache::QuerySysIconIndexForPath` (`IconCache.cpp:1148`) calls `SHGetFileInfoW` *without* `SHGFI_USEFILEATTRIBUTES`, statting the live path. Runs on the UI thread from `EnsureD2DResources` (`:378`), the deferred-init/paint path (`NavigationView.cpp:758`), and synchronously on every `SetPath` (`:1241`). Failures are deliberately not cached (`IconCache.cpp:1163`), so an offline UNC root re-incurs the full timeout on every navigation.
Scenario: navigating to `\\deadhost\share` freezes the pane (and pump, since it happens inside paint) for the SMB/DNS timeout.
Fix: resolve the icon off the UI thread (post to a worker, update `_menuIconBitmapD2D` via posted message, mirroring disk info), or at minimum pass `SHGFI_USEFILEATTRIBUTES` and skip live resolution for UNC / non-ready drive roots.

**Per-keystroke synchronous `LoadLibraryExW` + plugin `Initialize` on the UI thread when the instance-context segment changes**
`RedSalamander/NavigationView.Edit.cpp:1382/1429` — corrected to **robustness** (downgraded from dos-oom) — **CONFIRMED mechanism**
`UpdateEditSuggest` runs on the UI thread from `OnTextChanged` with no debounce (`:657-665`). In the `needsInstanceContext` branch it synchronously does `LoadLibraryExW` (`:1382`), `RedSalamanderCreate` (`:1403`), and `initializer->Initialize(...)` (`:1429`) before handing off to the worker; only the directory enumeration is offloaded. A `(shortId, instanceContext)` cache (`:1372`) means this fires only on cache-miss keystrokes that mutate the segment *before* the `|`, not literally every keystroke. Severity tempered because the only plugin implementing `IFileSystemInitialize` is FileSystem7z, whose `Initialize` is cheap and non-networked; the worst-case network-hang scenario is not currently reachable in this tree. Residual real cost: synchronous DLL load + COM construction per context-segment edit.
Fix: move the `LoadLibraryExW`/`RedSalamanderCreate`/`Initialize` sequence onto the edit-suggest worker (it already carries `keepAlive` across the boundary), or gate instance mounting behind an explicit commit/debounce. Never construct the instance on the UI thread.

**Breadcrumb layout: unbounded segment count, one DWrite layout per segment, O(n²) collapse search, all synchronous**
`RedSalamander/NavigationView.Breadcrumb.cpp:274` — dos-oom — **PLAUSIBLE**
`SplitPathComponents` (`:654`) emits one `PathSegment` per component with no cap; `UpdateBreadcrumbLayout` creates an `IDWriteTextLayout` per part (`:195-202`) and runs a nested `prefixCount × suffixCount` collapse search (`:274-304`). Synchronous on every path change (`NavigationView.cpp:1235`) and every resize/DPI relayout. Correction: the O(n²) loop body is O(1) arithmetic (prefix sums), so the real cost is the O(n) DWrite layout creations for thousands of *distinct* components (the 256-entry cache clears wholesale on overflow), not the collapse loop the finding emphasized.
Scenario: a pasted/plugin path normalizing to thousands of `/`-separated parts → thousands of layout allocations + measurement on the UI thread per relayout → multi-hundred-ms-to-second stall.
Fix: cap rendered segments (keep first N / last M, collapse the middle to an ellipsis *before* measuring) and/or replace the O(n²) collapse with a prefix-sum binary search.

**Full-path popup segment layout is unbounded in segment and line count**
`RedSalamander/NavigationView.FullPathPopup.cpp:878` — corrected to **robustness** — **PLAUSIBLE**
`BuildFullPathPopupLayout` loops every `SplitPathComponents` part with no cap, creating an `IDWriteTextLayout` per segment (`:885`) plus a binary-search `TruncateTextToWidth`; `UpdateFullPathPopupWindow` pre-measures every part again (`:609-611`). Same uncapped source as the breadcrumb. Downgraded because the popup opens only on an explicit gesture, the worst stall already happens at navigation time in the breadcrumb (not unique to the popup), and reaching thousands of file-plugin components requires adversarial/synthetic plugin input. Transient stall + bounded memory that recovers.
Fix: cap rendered segments (collapse middle to ellipsis) and bound the pre-measure loop before allocating per-segment layouts.

**Breadcrumb `Present` forwards an unvalidated (possibly inverted/negative) dirty rect → device-discard thrash**
`RedSalamander/NavigationView.Breadcrumb.cpp:50` — dos-oom — **PLAUSIBLE**
`OnSize` (`NavigationView.cpp:994`) computes `_sectionPathRect.right = cx - diskInfo - history` with no clamp; for `cx < 122` the rect inverts, for `cx < 94` it goes negative/out-of-bounds. `RenderPathSection` forwards it verbatim as `pDirtyRects` to `Present1` (`Rendering.cpp:988-994`). The dual-pane splitter has no minimum width (`FolderWindowInternal.h:112-113`), so a pane can be dragged this narrow. `DXGI_ERROR_INVALID_CALL` is classified as device-loss (`Rendering.cpp:37-39`), triggering `DiscardD2DResources` + full recreate, then re-presenting the same bad rect → thrash while held narrow. PLAUSIBLE not CONFIRMED because the in-bounds-but-inverted case depends on undocumented retail `Present1` behavior (the negative/out-of-bounds case does reject). Self-limited; recovers on widen.
Fix: normalize and clamp the dirty rect (skip present or full-present when `right<=left || bottom<=top`, intersect with back-buffer bounds), and clamp `_sectionPathRect.right >= left` in `OnSize`.

**Siblings dropdown builds an unbounded menu from a directory listing**
`RedSalamander/NavigationView.Menus.cpp:2514` — dos-oom — **CONFIRMED**
`TryGetSiblingFolders` (`:2364/2389-2407`) appends every directory entry with no cap (contrast edit-suggest's `kEditSuggestMaxCandidates=256` / `kEditSuggestMaxItems=11`). `ShowSiblingsDropdown` builds one `MenuFlyoutItem` per sibling and shows them in a single synchronous modal `ContextMenu::Show` on the UI thread; identical pattern in `FullPathPopup.cpp:447-466`.
Scenario: opening siblings on a folder under a parent with tens of thousands of subdirectories (a generated cache tree) → O(N) allocations and a multi-second build/render stall, unusable menu.
Fix: apply the edit-suggest bound (small max + explicit `…` overflow affordance) so menu size is independent of directory cardinality. (This is the same enumeration that also drives the ID-overflow contract bug below.)

---

### RACE / STALE-ASYNC (incorrect behavior)

**Escape dismisses the suggestion popup without invalidating the in-flight request → stale popup reappears**
`RedSalamander/NavigationView.Edit.cpp:776` — incorrect-behavior — **CONFIRMED**
The Escape handler (`:774-788`) calls `CloseEditSuggestPopup()` (which only resets the popup HWND, `:1686`) and returns without bumping `_editSuggestRequestId` or resetting `_editSuggestPendingQuery`. The worker already moved the query out under the lock and posts results with the still-current `requestId`; `OnEditSuggestResults` (`:370`) accepts it (id still matches), rebuilds items, and re-creates the popup because `_editMode` is still true. Contrast `ExitEditMode` (`:849-854`), which correctly `fetch_add`s the request id.
Scenario: Escape closes the dropdown; ~1-2s later (network enumeration completes) it silently pops back open over the edit field.
Fix: in the Escape-dismiss branch (and `CloseEditSuggestPopup` when used as an explicit dismissal), `fetch_add` `_editSuggestRequestId` and reset `_editSuggestPendingQuery` under `_editSuggestMutex`.

**`OnEditSuggestResults` reopens the popup without confirming the field text still matches**
`RedSalamander/NavigationView.Edit.cpp:364` — incorrect-behavior (finding's "wrong-target" framing **overstated/corrected**) — **CONFIRMED**
The handler validates *only* `requestId` (`:370`), not `_editMode`, `_pathEdit`, or text correspondence, then unconditionally assigns `_editSuggestItems` and calls `UpdateEditSuggestPopupWindow` (which guards only `_editMode`/`_pathEdit`, not text). Several dismissal paths close the popup without bumping the id — Escape (`:774`) and `ApplyEditSuggestIndex` (`:1934-1965`, which `SetText`s without firing `OnTextChanged` — verified `TextField::SetText` does not invoke the callback, `DxUi.TextInput.cpp:937`). So a late result re-shows a popup inconsistent with the current field. Correction: clicking a re-shown stale entry navigates to *that entry's own* `insertText`, so it is not literally wrong-target — it's a spuriously-reappearing/inconsistent popup.
Fix: have `OnEditSuggestResults` additionally confirm `_editMode && _pathEdit->field` and that the enumerated folder matches the current trimmed field text before applying; pair with bumping `requestId` on every dismissal.

---

### CONTRACT / ROBUSTNESS

**Sibling dropdown command IDs are unbounded and overflow the documented 600-699 range into the history range**
`RedSalamander/NavigationView.Menus.cpp:2520` and `RedSalamander/NavigationView.FullPathPopup.cpp:459` — contract — **CONFIRMED (latent; not a live misroute)**
`ID_SIBLING_BASE=600`, `ID_HISTORY_BASE=700` (`NavigationView.h:836-838`) leave exactly 100 slots, but the sibling loop assigns `ID_SIBLING_BASE + i` with no cap over the uncapped `TryGetSiblingFolders` list. A parent with >100 subfolders (e.g. `C:\Windows\WinSxS`, `node_modules`) produces IDs 700+. Unlike the nav-menu (`:1169`) and drive-menu (`:2294`) loops, which enforce `nextId > maxId` truncation, siblings have no guard. **No live bug today**: each `ContextMenu::Show` returns its own id synchronously (`DxUi.h:278`), consumed in-scope and re-bounded against `_navDropdownPaths.size()` (`:2613-2618`); the `OnCommand` sibling branch (`NavigationView.cpp:1065`) is a documented no-op; sibling and history menus never coexist. The hazard is for future shared-range routing and the wrong header allocation map.
Fix: add `ID_SIBLING_MAX=699`, cap the sibling loop (mirroring `kMaxActions`), and tighten the consumer to require `selectedId <= ID_SIBLING_MAX` in both `ShowSiblingsDropdown` and `ShowFullPathPopupSiblingsDropdown`.

**Full-path popup edit validates the raw typed string but navigates the env-expanded path → rejects valid `%VAR%` paths**
`RedSalamander/NavigationView.FullPathPopup.cpp:1185` — incorrect-behavior — **CONFIRMED**
`ExitFullPathPopupEditMode` reads `buffer = field->GetText()` (`:1183`), calls `ValidatePath(buffer)` on the *unexpanded* text (`:1185`), and only afterward builds the target via `NormalizeUserTypedLocationText(buffer)` (`:1187`, which is what expands `%VAR%`). `ValidatePath` on the literal `C:\Users\%USERNAME%\Desktop` fails `GetAttributes` (the literal doesn't exist) → rejected with `IDS_FMT_INVALID_PATH` (`:1235`). The inline breadcrumb editor does the correct opposite order (`NavigationView.Edit.cpp:861-863`: normalize first, then validate), so identical input succeeds inline but fails in the popup.
Fix: expand/normalize first, then validate the normalized string, matching `NavigationView.Edit.cpp:861-863`.

**Full-path popup `WM_ACTIVATE` destroys the popup while a child dropdown is open → menu-session corruption**
`RedSalamander/NavigationView.FullPathPopup.cpp:298` — robustness (downgraded from the finding's UAF framing) — **CONFIRMED mechanism**
Opening a siblings dropdown calls `ContextMenu::Show(popupHwnd, ...)` (`:498`); inside, the menu's `SetActiveWindow` on its own (owned, non-child, `WS_EX_TOOLWINDOW`) root sends `WA_INACTIVE` to the popup. `OnFullPathPopupActivate` (`:298-303`) has no guard for an open dropdown and `IsChild` is false for an owned window, so it falls through to `CloseFullPathPopup` → `_fullPathPopup.reset()` (`:788`), destroying the popup (and its owned menu root) *while the modal `Show` loop is still on the stack*. Trigger is deterministic on every dropdown open, not just external activation steals. The literal use-after-free is mitigated by null guards in `RenderFullPathPopup` (`:936/947`), so the concrete outcome is menu-session corruption / popup self-destruction, not a guaranteed UAF. Main breadcrumb dropdowns avoid this by anchoring on `GetAncestor(GA_ROOT)` (`Menus.cpp:1416` etc.); the popup uniquely anchors on itself.
Fix: suppress the `WA_INACTIVE` close while a popup dropdown is active (guard on `_fullPathPopupActiveSeparatorIndex != -1`), or anchor the sibling `ContextMenu` on the root owner window so the popup never deactivates underneath an open menu.

**Change Drive submenu assigns a live command id to command-only items but registers no action → dead no-op entries**
`RedSalamander/NavigationView.Menus.cpp:1376` — robustness (downgraded; not currently reachable) — **CONFIRMED mechanism, unreachable today**
In the Change Drive loop an item is actionable when `hasPath || hasCommand`, consumes an id via `child.commandId = fileId++` (`:1378`), but only pushes a `MenuAction` when `hasPath` (`:1381-1386`). A command-only item gets a live id with no `_navigationMenuActions` entry, so `ExecuteNavigationMenuAction` (`:752-786`) returns false → silent no-op. Sibling builders handle the command case via an else branch (`:1188-1197`, `:1567-1576`). **Not reachable today**: the file plugin's `GetMenuItems` (`FileSystem.Menu.cpp:184-310`) emits only path-bearing entries (`commandId` defaults to 0), so `hasCommand && !hasPath` is impossible; would only manifest if a future file plugin returns a command-only nav item.
Fix: register a `MenuAction` for the command case too (mirroring the sibling builders), or treat command-only items as non-actionable so no id is consumed.

**History dropdown snapshot can show a stale checkmark/ordering after async history mutation during the modal menu**
`RedSalamander/NavigationView.Menus.cpp:1984` — robustness (downgraded from incorrect-behavior) — **PLAUSIBLE, cosmetic only**
`ShowHistoryDropdown` snapshots `_pathHistory` into `_navDropdownPaths` and computes `checked` at build time (`:2038-2042`). The modal `ContextMenu::Show` pump (`:2092`) dispatches foreign messages, and an external folder relocate can drive `SetHistory` (`NavigationView.cpp:1290`), mutating `_pathHistory`/`_currentPath` without touching the snapshot. **No crash, no wrong navigation**: the snapshot is bounds-checked (`:2114`) and the user navigates to exactly the path they clicked. Only the build-time checkmark and ordering can be momentarily stale; both reset on reopen. Requires an external process to relocate the *current* folder during the narrow menu-open window.
Fix (low priority): recompute the checked entry from `_currentPath` at show time, or gate history re-render behind a "dropdown open" flag and rebuild on close.

**`RenderDiskInfoSection` draws the history chevron into the History rect (copy-paste artifact)**
`RedSalamander/NavigationView.Rendering.cpp:779` — robustness (downgraded; no visible artifact) — **CONFIRMED dead code, REFUTED visible bug**
Lines `776-798` are a verbatim copy of the history-chevron block from `RenderHistorySection`, drawing into `_sectionHistoryRect` — a region this function doesn't own. But the `endDraw` present uses only `dirtyRect = _sectionDiskInfoRect` (`:767`) with `DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL` dirty-rect present, so those pixels lie outside the presented region and never reach the screen; the glyph/rect/brush are also identical to the authoritative draw. So it is wrong-region *dead work*, not a visible artifact.
Fix: remove the history-chevron block from `RenderDiskInfoSection`; it should render only within `_sectionDiskInfoRect`.

---

## 3. Cross-Cutting Themes

1. **UI-thread blocking I/O is endemic and the fix is already in the codebase.** Five confirmed entry points (commit validation ×2, drive-menu enumeration, menu-icon resolution, plugin mount) plus two unbounded-layout paths all call blocking filesystem/shell/loader APIs synchronously on the UI thread. The project already ships the correct pattern — `jthread` workers + `stop_token` + `PostMessagePayload` (edit-suggest worker, `_driveInfoCv` worker). The defect is consistently *not reaching for the existing async primitive on the commit/menu/render paths*. A single policy ("no blocking FS/shell/loader call on the UI thread; route through a worker with a bounded timeout") closes the entire cluster. This directly violates the explicit CLAUDE.md no-UI-thread-blocking rule.

2. **Validate-vs-navigate operates on divergent path strings.** `ValidatePath` and the actual navigation repeatedly disagree about *which* string they're checking: drive-relative input is resolved against process CWD, env-var expansion happens before validation in one path and after in another, and the breadcrumb vs. popup commit do expand/validate in opposite orders. The root remedy is a single canonical "normalize → validate the normalized string → navigate the same string" pipeline shared by both editors, with explicit handling for drive-relative and literal-`%`-token inputs.

3. **Async results and posted requests outlive their context without a generation guard.** Autocomplete results, the suggestion popup, and posted plugin-navigate requests all survive dismissal/plugin-swap because the consumer checks too little (only `requestId`, or nothing). The unifying fix is a strict generation/identity guard bumped on *every* dismissal and re-checked (plus identity captured in posted payloads) before any state mutation or navigation.

4. **Menu-ID ranges are a fragile manual contract.** Sibling IDs overflow their documented 100-slot range; only the per-menu in-scope `Show`-return consumption pattern saves it from being a live misroute. The header allocation map is already wrong. Dispatch-by-stored-index (rather than reconstructing an index from a global ID offset) would eliminate the whole class.

---

## 4. Coverage & Residual Risk

**Covered:** address-bar path input parse/validate/navigate (drive-relative, UNC, env-var); inline + full-path-popup commit paths; autocomplete async lifecycle (request-id staleness, popup re-show); per-keystroke plugin mount; breadcrumb segment layout and dirty-rect present; sibling/history/drive dropdown enumeration, ID allocation, and lifetime; menu-icon and disk-info rendering; core path-sync and posted plugin-navigate replay.

**Confirmed clean / refuted (don't re-litigate):**
- **No code-execution / injection surface.** The address bar performs navigation only; `%VAR%` expansion is pure process-local string substitution — the "host-controlled target / injection" framing was explicitly refuted.
- **No confirmed memory-safety defect.** The popup-teardown "UAF" degraded to menu-session corruption guarded by null checks; posted-navigate payloads are RAII-owned (no UAF/double-free).
- **Sibling-ID overflow is latent only** — no live misroute under the current in-scope `Show`-return dispatch model.
- The `RenderDiskInfoSection` chevron produces **no visible artifact** (dirty-rect present discards it).

**Residual risk / not fully verified:**
- **Breadcrumb dirty-rect thrash** (`Breadcrumb.cpp:50`) and the **in-bounds-inverted-rect** behavior of retail `Present1` were not runtime-confirmed — needs observation at a narrow splitter width (the negative/out-of-bounds sub-case is sound).
- **Per-keystroke plugin mount** severity depends on a future networked `IFileSystemInitialize` implementation; only FileSystem7z implements it today (cheap, non-networked), so the network-hang scenario is currently unreachable but would regress silently if a Curl/SFTP/S3/Drive plugin adds `Initialize`.
- **Plugin-navigate replay after `SetFileSystem`** and the **history-snapshot staleness** both require a specific UI-thread interleaving (queued message + concurrent swap/relocate) — plausible but not deterministically demonstrated; consequences are bounded (mis-navigation within the app's namespace; cosmetic checkmark) with no memory-safety angle.
- **Not audited here:** keyboard/accessibility navigation of the breadcrumb and menus, theming/DPI correctness of the rendered segments beyond the dirty-rect path, and clipboard/paste sanitization beyond path parsing. The full-path-popup *siblings* path is also noted as lacking self-test coverage (only edit-route and ancestor-click are tested), consistent with the teardown bug going unnoticed.
