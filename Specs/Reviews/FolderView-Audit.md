> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


The references check out (CF_HDROP walk at lines 532-537 with no `GlobalSize`; `_currentFolder.value()` at 572; clipboard `*effect` at 267 with no size check). Here is the final report.

---

# FolderView Audit — Final Report

## 1. Verdict

FolderView's drag-drop and file-op entry points are **not safe to ship as-is.** The audit surfaced a tight cluster of confirmed, high-impact defects whose common denominator is that **FolderView trusts inputs and timing it does not control** — the OLE data object, the clipboard, the drop coordinate, the OS move-completion handshake, and the async enumeration generation — and acts destructively on them. The most dangerous themes, in order:

- **Move-onto-self destroys directories.** A Shift-drag of a *directory* back into its own folder routes a move-onto-self all the way through the FS plugin, where the copy-then-delete fallback deletes the source after copying nothing. Permanent loss, no prompt. (Confirmed.)
- **The drop point is ignored — every drop lands in `_currentFolder`.** Dropping onto a visible subfolder row silently copies/moves into the *parent* instead. With Shift this is a wrong-target MOVE. This single root cause feeds four findings.
- **MOVE reported to Explorer before the async copy runs.** We echo `DROPEFFECT_MOVE` back to the OLE source on a queue-time `S_OK`; Explorer can delete the source before our queued copy to a slow/cloud/archive destination has (or hasn't) succeeded.
- **Unvalidated OLE/clipboard parsing.** The CF_HDROP fallback and the internal-drop `reserve` walk attacker-controlled buffers with no `GlobalSize` bound → OOB read / `std::terminate`. Any process can register the internal clipboard format.
- **Selection model lets the wrong files get dragged.** A plain click-drag of an *unselected* item, or a drag started on empty background, carries the *previous* selection / focused item — a wrong-target move of files the user never picked.
- **Async-enumeration data race on `_itemsFolder`.** The icon/thumbnail worker reads a `std::filesystem::path` the UI thread move-assigns under no lock — torn read / read of freed buffer during ordinary navigation.

Below, findings are grouped by severity with same-root-cause defects deduplicated. Every confirmed data-loss / wrong-target / crash / security finding is retained.

---

## 2. Findings

### Tier 1 — Data-loss / Wrong-target / Crash / Security / UAF

---

**D1. Directory move-onto-self is routed into the FS and destroys the source (no self-drop guard)**
`RedSalamander/FolderView.DragDrop.cpp:550` (root cause; also dispatched at `:572`, `:607`)
Severity: **data-loss** — CONFIRMED.
Gesture: Shift-drag a selection containing a directory and drop it onto its own pane background (`_currentFolder` is already the items' parent). `ResolveDropEffect` forces MOVE; `PerformDrop` issues `FILESYSTEM_MOVE` with destination == source-parent and never compares source parent against destination.
Timing/failure: For a *directory* self-drop, `directoryMergeDestination` is true, the rename fallback leaves children behind (`FileSystem.FileOps.cpp:5067/5078` returns `S_OK` before delete, `fullyMoved=false`), control falls to copy+delete (`:5356`), the copy silently no-ops (same dir treated as merge target), then `DeleteCopiedSourceEntryForMove` (`:5467`) deletes every source file because `DestinationMatchesSourceFile` is true for a file vs itself, and finally removes the emptied source directory. The dropped directory and all contents are permanently destroyed; nothing is copied anywhere; no conflict prompt fires. (Plain *file* self-drop is a harmless `ERROR_ALREADY_EXISTS` no-op.)
Fix: Reject in the drop channel before dispatch. For each source, normalize and case-insensitively compare `parent_path()` against `_currentFolder`; drop same-folder MOVE entries (or cancel the gesture). This guard belongs in `PerformDrop`, mirroring the file-ops bridge's `isSameOrUnderFolder` logic — do not rely on the FS plugin, which has no same-path early-out.

---

**D2. Drop ignores the drop point — files always land in `_currentFolder`, never the hovered subfolder** *(root cause for 4 findings)*
`RedSalamander/FolderView.DragDrop.cpp:572` (also `:603/:607`; `Drop` receives `POINTL` at `:126` but forwards it only to the visual helper at `:142`)
Severity: **wrong-target** — CONFIRMED.
This single defect — `PerformDrop` has no `point` parameter and no hit-test against `_items` — manifests as four reported symptoms; they share one fix:
- Dropping onto a visible subfolder row copies/moves into the current (parent) folder instead of *into* that subfolder. With Shift, a silent wrong-target MOVE relocates the files and the user must hunt for them. (`:572`)
- No-modifier intra-app drag onto a subfolder row silently COPIES the selection into its own parent folder, producing "Copy of…" duplicates rather than moving into the target. (`:572`, dup of the above)
- No same-folder / drop-onto-dragged-item / descendant rejection anywhere in `PerformDrop` (early-outs are only `!_currentFolder`, `effect==NONE`, `paths.empty()`).
Failure: User releases over folder `Archive`; 50 files land beside `Archive` in the current folder, not inside it. `DragOver` also never highlights a target, so there is no visual cue.
Fix: Plumb `point` into `PerformDrop`. Hit-test against `_items`; if it lands on a directory item, use that item's full path (`GetItemFullPath`) as the destination, else `_currentFolder`. Reject (`DROPEFFECT_NONE`) when the resolved destination equals a source's parent, equals a dragged item, or is a descendant of a dragged directory. Reflect the resolved target in `DragOver` so the cursor matches what `Drop` will do.

---

**D3. Descendant-move data-loss path (drop a folder onto a pane showing its own child)**
`RedSalamander/FolderView.DragDrop.cpp:549`
Severity: **data-loss** — PLAUSIBLE (constructible; primary headline scenario is OS-blocked).
Mechanism: `PerformDrop` → `StartFileOperationFromFolderView` → `FileSystem::MoveItems` never excludes the destination subtree from enumeration. The common case (drag `C:\Work` onto a pane showing `C:\Work\Sub`, dest resolves to non-existent `C:\Work\Sub\Work`) is blocked by the OS: `MoveFileWithProgressW` fails `ACCESS_DENIED`, the copy+delete fallback is never reached (`FileSystem.FileOps.cpp:5321-5325`). **The real loss path requires the descendant destination to already exist as a directory:** then `directoryMergeDestination=true` → `MoveDirectoryMergeByRename` fails → `CopyPathInternalWithDirectoryParallelism` recurses into a growing tree with no destination-subtree exclusion, followed by source delete.
Fix: Same UI-layer guard as D2 — reject when the destination is a descendant of any dragged directory (case-insensitive prefix check with a path-separator boundary). The FS-layer enumeration should also exclude the destination subtree as defense-in-depth.

---

**D4. MOVE echoed to the external OLE source before the async copy runs**
`RedSalamander/FolderView.DragDrop.cpp:575` (reported as `:693` `*performedEffect = effect`; back to OLE at `:151`)
Severity: **data-loss** — PLAUSIBLE (RedSalamander-side defect confirmed; final delete is in the source app).
Mechanism: For an external Explorer MOVE drop, `_fileOperationRequestCallback` → `StartOperation` spawns a detached worker (`FolderWindow.FileOperations.State.Runtime.Part.cpp:759`) and returns `S_OK` at *queue* time. `PerformDrop` treats `S_OK` as success and reports `DROPEFFECT_MOVE` as performed. For classic CF_HDROP filesystem drags Explorer deletes the source on a reported logical MOVE.
Failure: Explorer deletes the source files immediately while our queued copy into a possibly-virtual/cloud/archive destination is still pending; if that copy later fails (disk full, permission, network drop, cancel), the data is unrecoverably lost.
Fix: Never echo MOVE based on a queue-time `S_OK`. Either return `DROPEFFECT_COPY` to the OLE source (and let our own verified move delete the source), or confirm completion before reporting the effect. Honor the `CFSTR_PERFORMEDDROPEFFECT` handshake.

---

**D5. Click-drag of an unselected item moves the *previously-selected* items**
`RedSalamander/FolderView.Interaction.cpp:477`
Severity: **wrong-target** — CONFIRMED.
Gesture: Select A,B,C (marquee/Ctrl), then plain left-click (no modifier) an unselected item Z and drag it to another pane.
Mechanism: The no-modifier branch calls `FocusItem(*hit,false)`, which only sets the `focused` flag and never collapses selection; `OnLButtonUp` does no deferred collapse; `SelectSingle` is reachable only from keyboard nav. When the drag threshold trips, `BeginDragDrop` → `GetSelectedOrFocusedPaths` returns the *selected* A,B,C.
Failure: The drag carries A,B,C, not the Z the user is visually dragging; a MOVE drop silently relocates the wrong files.
Fix: On a no-modifier click that hits an *unselected* item, collapse selection to that item (`SelectSingle(*hit)` in the no-modifier branch, or a deferred `SelectSingle` in `OnLButtonUp` when no drag occurred). Preserve the multi-selection for a drag only when the clicked item is itself already selected.

---

**D6. Drag started on empty background drags the previously-focused item**
`RedSalamander/FolderView.Interaction.cpp:483`
Severity: **wrong-target** — CONFIRMED. (Same family as D5: stale source identity feeding `GetSelectedOrFocusedPaths`.)
Gesture: Left-press empty space (clears selection, leaves `_focusedIndex` set), then drag past the threshold intending a marquee select.
Mechanism: `OnLButtonDown` unconditionally `SetCapture` + `_drag.dragging=true`; the no-hit else branch calls `ClearSelection()` but does not clear `_focusedIndex` or disarm the drag (`hasStartItemRect=false`, so the "stay inside start item" guard is skipped). `BeginDragDrop` → `GetSelectedOrFocusedPaths` falls back to `_items[_focusedIndex]`. Marquee is unimplemented (`Specs/UI/UI_FolderView.md:523`), so the empty-area drag becomes an OLE file drag.
Failure: User clicks empty space to deselect then drags to rubber-band; instead an OLE drag of the previously-focused file starts, MOVE-able on drop.
Fix: In the no-hit else branch set `_drag.dragging=false` (or don't `SetCapture`/arm unless `hit` has a value); alternatively gate `BeginDragDrop` on `_drag.anchorIndex != (size_t)-1`. Route empty-space drags to selection once marquee lands.

---

**D7. Drop/paste destination uses in-flight `_currentFolder` while the pane still shows `_displayedFolder`**
`RedSalamander/FolderView.DragDrop.cpp:572` (and paste at `RedSalamander/FolderView.FileOps.cpp:756`)
Severity: **wrong-target** — PLAUSIBLE (timing-dependent).
Mechanism: `SetFolderPath` sets `_currentFolder` synchronously and starts async enumeration; `_displayedFolder`/`_itemsFolder` update only on completion (`FolderView.Enumeration.cpp:1664/1672`). `PerformDrop` and `PasteItemsFromClipboard` resolve the destination from `_currentFolder` with no `_currentFolder == _displayedFolder` check. A guard (`IsCurrentFolderEnumerated()`, `FolderView.cpp:291`) already exists but is not called on this path.
Failure: During slow/network navigation A→B, a drop/paste while A's contents are still visible lands in B; a MOVE silently relocates files to the wrong folder.
Fix: Gate drop/paste on `IsCurrentFolderEnumerated()`, or resolve the destination from `_displayedFolder`/`_itemsFolder` (what the user sees), deferring or rejecting otherwise.

---

**D8. Context-menu command targets the wrong file after an in-loop async refresh migrates focus**
`RedSalamander/FolderView.Menus.cpp:367` / `:372` *(two findings, one root cause)*
Severity: **wrong-target → data-loss** — PLAUSIBLE (timing race).
Mechanism: A single right-click only *focuses* (`FocusItem(*hit,false)`, `:354`); `ContextMenu::Show` runs a nested modal loop that dispatches posted `kFolderViewEnumerateComplete`/`kFolderViewDirectoryCacheDirty` to `WndProc`. If an external process renames/deletes the right-clicked item while the menu is open, name-based focus restoration misses and `ApplyCurrentSort` falls back to a raw positional index (`Enumeration.cpp:1318-1332`) pointing at a *different* file. The posted `WM_COMMAND` then resolves the target via the corrupted `_focusedIndex` — and `CommandDelete` recycles with **no confirmation dialog** (`FolderWindow.FileOperations.cpp:1057/1079`).
Failure: User right-clicks A → Delete; A was just removed externally; focus falls onto neighbor B; B is recycled, A untouched. Loss of an item the user never selected.
Fix: Snapshot the intended target by display name at menu-build time and re-validate after `Show` before posting (or pass it explicitly into the command). In `ApplyCurrentSort`, do **not** fall back to a raw positional index when the named focus target disappeared during a refresh — clear focus instead so the destructive command becomes a no-op. Best: suppress directory-refresh `_items` reassignment while a modal menu / `DoDragDrop` loop is active.

---

**S1. CF_HDROP / DROPFILES path-list walk has no `GlobalSize` bound (OOB read)** *(three findings, one root cause)*
`RedSalamander/FolderView.DragDrop.cpp:532`
Severity: **security / crash** — CONFIRMED.
Mechanism: The CF_HDROP fallback (external/foreign drop) `GlobalLock`s the medium and walks `current = base + dropFiles->pFiles` with `while (current && *current) { paths.emplace_back(current); current += wcslen(current)+1; }`. No `GlobalSize` is obtained; `pFiles` (attacker-controlled offset) is never validated against `sizeof(DROPFILES)` or the allocation; `dropFiles->fWide` is dereferenced with no `size >= sizeof(DROPFILES)` guard. This directly contrasts the internal-format branch above it, which computes `bytesAvailable = GlobalSize` (`:405`) and bounds every read.
Failure: A short HGLOBAL, an out-of-range `pFiles`, or a missing terminating double-NUL within the allocation makes `wcslen`/`*current` read past the mapped buffer → access violation (crash), or adjacent heap bytes become bogus `std::filesystem::path` operands handed to a real MOVE/COPY (wrong-target). Every external Explorer drag hits this branch.
Fix: Capture `bytes = GlobalSize(medium.hGlobal)`; reject `bytes < sizeof(DROPFILES)` and `pFiles < sizeof(DROPFILES) || pFiles >= bytes`; bound the walk to `[base+pFiles, base+bytes)`, requiring the terminator inside the allocation. Prefer `DragQueryFileW` (length-safe), as `ReadFileDropClipboard` already does (`FolderView.FileOps.cpp:220`).

---

**S2. Attacker-controlled `pathCount` drives `vector::reserve` inside a `noexcept` lambda → `std::terminate` (DoS)**
`RedSalamander/FolderView.DragDrop.cpp:460`
Severity: **crash** — CONFIRMED.
Mechanism: `tryReadInternalDrop` is `noexcept` (`:377`); the only size gate is `bytesAvailable >= 16` (`:406`) and only `version` is checked (`:427`). At `:460` `paths.reserve(header->pathCount)` runs *before* the per-entry bounds checks. The internal clipboard format is registered purely by name (`FolderViewInternal.h:1128`), so any process can supply it.
Failure: A 16-byte buffer with `version=1`, zero-length strings, `pathCount=0xFFFFFFFF` makes `reserve` request ~171 GB → `std::length_error`/`std::bad_alloc` crosses the `noexcept` boundary → `std::terminate` kills the whole file manager.
Fix: Bound `pathCount` before reserving (reject `> bytesAvailable / sizeof(uint32_t)`), or drop the `reserve` and let the bounds-checked loop grow the vector. Never `reserve` an attacker-controlled count inside `noexcept`.

---

**S3. `ReadPreferredDropEffectClipboard` dereferences a clipboard `DWORD` with no size check (OOB read)**
`RedSalamander/FolderView.FileOps.cpp:267`
Severity: **security / robustness** — CONFIRMED (catastrophic-MOVE framing overstated).
Mechanism: `effect = GlobalLock(handle); result = *effect;` with no `GlobalSize(handle) >= sizeof(DWORD)` check. A clipboard owner controls the HGLOBAL size for `CFSTR_PREFERREDDROPEFFECT`.
Failure: A zero-/sub-4-byte buffer makes `*effect` an OOB read (usually within the granularity-rounded heap block, so a crash is unlikely but the read is still out of the requested size). If the garbage equals `DROPEFFECT_MOVE`, `PasteItemsFromClipboard` (`:739`) performs a destructive MOVE instead of COPY. (Note: a real attacker wanting MOVE would just write a valid `DWORD=2`, so the data-loss angle is weak — the OOB read is the real defect.)
Fix: After `GlobalLock`, verify `GlobalSize(handle) >= sizeof(DWORD)`; on failure return `std::nullopt` (caller already defaults to COPY via `value_or`).

---

### Tier 2 — Race / Leak

---

**R1. Icon/thumbnail worker reads non-atomic `_itemsFolder` while the UI thread move-assigns it (data race / torn read of freed buffer)** *(two sites, one root cause)*
`RedSalamander/FolderView.Icons.cpp:539` (icon queue) and `:911` (thumbnail queue)
Severity: **race** — CONFIRMED.
Mechanism: `ProcessIconLoadQueue`/`ProcessThumbnailLoadQueue` run on the background worker and execute `perf.SetDetail(_itemsFolder.native())` *before* taking `_enumerationMutex` (`:552`/`:924`). `_itemsFolder` is a plain `std::filesystem::path` (`FolderView.h:919`). The UI thread does `_itemsFolder = std::move(payload->folder)` (`Enumeration.cpp:1664`) and `.clear()` (`FolderView.cpp:522`) with no lock. `SetDetail` copies the path's internal wstring buffer by view.
Timing: User navigates A→B (or a watcher refresh fires) while the worker is draining A's queue; the worker's `.native()` read races the UI thread's move-assign of the same wstring → torn perf-detail string at best, read of a freed/reallocated heap buffer at worst. `CancelPendingEnumeration` clears queues but never joins/pauses the worker, so it can already be inside the function past the generation check.
Fix: Stop reading `_itemsFolder` off the UI thread. Use the per-request owned `fullPath` for perf detail, or snapshot the folder path into the request under `_enumerationMutex` at queue time. Never touch `_itemsFolder` across the UI/worker boundary without synchronization.

---

### Tier 3 — Incorrect-behavior / Contract / Robustness

---

**B1. UI thread blocks on a synchronous `IShellItemImageFactory::GetImage` for offline/network thumbnails**
`RedSalamander/FolderView.Icons.cpp:1004` (flags at `:98`)
Severity: **incorrect-behavior** (UI-thread block) — CONFIRMED.
Mechanism: `ExtractShellThumbnailBitmap` calls `GetImage` with `SIIGBF_THUMBNAILONLY|BIGGERSIZEOK|SCALEUP` — **no `SIIGBF_INCACHEONLY`** — so it performs synchronous disk/network I/O and runs the in-proc `IThumbnailProvider`, blocking for the full SMB/handler timeout. Cancellation signals are checked only between items; teardown `StopEnumerationThread` does `request_stop()` then joins the worker on the UI thread (`Enumeration.cpp:235/245`), and `request_stop()` cannot interrupt a synchronous shell call.
Failure: On a slow/disconnected share in Thumbnails mode, navigation/close freezes for the timeout. Violates the CLAUDE.md no-UI-thread-blocking rule.
Fix: Add `SIIGBF_INCACHEONLY` (or a bounded fast-first pass) and a watchdog/deadline; make teardown not block on the worker (detach or deadline the in-flight extraction).

---

**B2. Thumbnails for virtual-FS (archive/cloud/curl) items parse a fabricated local-looking path against the real shell namespace**
`RedSalamander/FolderView.Icons.cpp:794` (shell parse at `:89`, WIC at `:145`)
Severity: **wrong-target** — CONFIRMED.
Mechanism: `QueueThumbnailLoading` gates only on `_thumbnailsVisible`/`_items`/`_hWnd`, never on `_fileSystem` type; `request.fullPath = GetItemFullPath(item) = _itemsFolder / displayName` is passed verbatim to `SHCreateItemFromParsingName` and `CreateDecoderFromFilename`. `IsFilePluginShortId` (already computed in `SetFileSystem`, `FolderView.cpp:556`) is not applied here. `FileSystem7z` reports `shortId="7z"` so its paths are virtual.
Failure: For `C:\archive.7z\sub\photo.jpg`, the shell/WIC parse the synthetic path against the real namespace. If a same-named real file resolves, the user sees a *different* local file's thumbnail (silent misrepresentation); otherwise wasted worker slots and unintended local-path probing for cloud/network items.
Fix: Gate `QueueThumbnailLoading`/`QueueIconLoading` extraction on a local-shell-backed capability flag. For virtual FS, route through the plugin's own extraction API or fall back to type icons.

---

**B3. WIC thumbnail decode runs an untrusted-image parser in-process with no source-dimension cap**
`RedSalamander/FolderView.Icons.cpp:145`
Severity: **robustness / security hardening** — PLAUSIBLE.
Mechanism: `DecodeWicThumbnailPixels` drives `CreateDecoderFromFilename`→`GetFrame`/`GetSize`/`CopyPixels` over attacker-controlled files. Source dimensions are only checked non-zero (`:161`); the *output* is clamped to 512px but the decoder still processes the full source frame (incl. third-party HEIC/WebP codecs). COM pointers are correctly `wil::com_ptr`-wrapped, so this is not a lifetime bug.
Failure: Browsing a folder with a malicious image silently invokes the codec on the worker; a codec memory-safety bug becomes folder-view-triggerable, and a decompression-bomb source frame spikes memory/CPU before the output clamp applies.
Fix: Before decoding, reject sources whose `width*height` exceeds a sane cap; prefer the shell cache path; consider isolated/limited decode for untrusted input.

---

**B4. Unbounded `reserve` driven by untrusted virtual-FS `GetCount` → `bad_alloc` → thread `terminate` (DoS)**
`RedSalamander/FolderView.Enumeration.cpp:393` (count at `:381`)
Severity: **robustness** — PLAUSIBLE (outcome real; finding misreads the catch site).
Mechanism: `files.reserve(std::max(estimatedFiles, 256u))` / `directories.reserve(...)` use `entryCount` from an untrusted plugin with no clamp, *before* the buffer-size sanity checks. **Correction to the original finding:** these reserves are *outside* the `try` block (`try` starts at `:396`), so the `catch(bad_alloc){ std::terminate(); }` at `:892` does NOT fire — the exception escapes `ExecuteEnumeration` to `EnumerationWorker` (`:317`) which has no try/catch, and an exception escaping a thread function still ends in `std::terminate`. In-tree plugins tie count to a materialized buffer; the exploitable case is a hostile/buggy third-party plugin returning an arbitrary cheap count across the COM boundary.
Fix: Clamp `entryCount` against `GetBufferSize()/sizeof(FileInfo)` and an absolute ceiling before reserving; treat `GetCount` as a hint. Optionally wrap the worker body so a thrown exception degrades gracefully.

---

**B5. External local-file drop onto a virtual-FS pane routes the destination FS as the *source* reader**
`RedSalamander/FolderView.DragDrop.cpp:564`
Severity: **incorrect-behavior** — CONFIRMED.
Mechanism: For an external CF_HDROP drop, `tryReadInternalDrop` returns `S_FALSE`, so `sourceContextSpecified=false`; in `StartFileOperationFromFolderView` (`FolderWindow.FileOperations.cpp:386`) the `if (isCopyMove && sourceContextSpecified)` block is skipped, leaving `fileSystem = destinationState.fileSystem` (the virtual FS) as the source reader for local Explorer paths and `destinationFileSystem` null.
Failure: Dropping `C:\Users\me\report.pdf` onto a 7z/S3/Drive pane asks the archive/cloud provider to open that local path inside its own namespace — read fails (or wrong object). The intuitive "add local files to the archive / upload to cloud" import never happens.
Fix: Route external CF_HDROP drops as an explicit local→destination cross-FS import (local source handle + pane FS as destination), or reject (`DROPEFFECT_NONE`) external drops onto panes that cannot import local paths.

---

**B6. No-modifier drop effect defaults to COPY regardless of volume (deviates from same-drive MOVE convention)**
`RedSalamander/FolderView.DragDrop.cpp:311`
Severity: **incorrect-behavior** (UX/convention; not data-loss) — CONFIRMED.
Mechanism: `ResolveDropEffect` derives the effect purely from modifier keys; with none held it returns COPY whenever copy is allowed (`:311-314`). No source/destination volume comparison exists anywhere in the drag-drop source.
Failure: A no-modifier same-drive drag that the user (per Windows convention) expects to MOVE instead COPIES, leaving a duplicate at the source. Internally consistent (cursor and action both COPY), so the failure direction is safe — but the common case is wrong.
Fix: With no modifier, compare source vs destination volume — same volume → MOVE if allowed, different → COPY — before the current copy-first fallback.

---

**B7. Device-lost (`D2DERR_RECREATE_TARGET` / `DXGI_ERROR_DEVICE_REMOVED`) never recreates the device; pane stays permanently blank**
`RedSalamander/FolderView.Rendering.cpp:1876` (also Present paths `:1912`, `:1936`)
Severity: **robustness** — CONFIRMED.
Mechanism: `EndDraw`/`Present` failure branches do `ReportError; ReleaseSwapChain; EnsureSwapChain;` without inspecting the HRESULT for device-lost and without calling `DiscardDeviceResources()`. `EnsureDeviceResources` early-returns while the (now-dead) device pointers are non-null (`:165`); `EnsureSwapChain` early-returns unless `_d3dDevice` is null (`:439`). `DiscardDeviceResources` (the only full rebuild) is called only from teardown (`FolderView.cpp:202/837`).
Failure: After a driver TDR / GPU hot-swap / dGPU↔iGPU switch, the pane goes blank and never repaints — every `Render` reuses the dead device.
Fix: In the `EndDraw`/`Present` failure branches, detect device-lost HRESULTs (and query `GetDeviceRemovedReason`), call `DiscardDeviceResources()`, then `InvalidateRect` to force a rebuild on the next `Render`.

---

**B8. `DrawItem` indexes `_items[_hoveredIndex]` guarded only by the sentinel, not by bounds**
`RedSalamander/FolderView.Rendering.cpp:2176`
Severity: **robustness / defense-in-depth** — PLAUSIBLE (latent; not currently reachable).
Mechanism: This is the only `_hoveredIndex` read that indexes the vector without a `_hoveredIndex < _items.size()` check; the three sibling sites (`Interaction.cpp:55/556/566`) all bounds-check. Today every `_items`-shrinking path resets `_hoveredIndex` on the UI thread before a render, so no live gesture reaches OOB.
Fix: Match the siblings: `_hoveredIndex != (size_t)-1 && _hoveredIndex < _items.size() && std::addressof(item) == std::addressof(_items[_hoveredIndex])`.

---

**B9. Per-item selected-text / hover brushes recreated every `DrawItem` in the hot loop**
`RedSalamander/FolderView.Rendering.cpp:2394` (hover at `:2252`)
Severity: **robustness / efficiency** (not a leak — `wil::com_ptr`) — CONFIRMED.
Mechanism: `CreateSolidColorBrush` per selected/hovered item per frame, vs the cached-member-`SetColor` pattern used in the same function for `_selectionBrush`/`_focusBrush`.
Failure: Scrolling a large folder with a wide selection causes hundreds of `ID2D1SolidColorBrush` creations per frame on the UI thread → avoidable jank.
Fix: Add cached `_selectedTextBrush`/`_hoverBrush` in `RecreateThemeBrushes`, reset in `DiscardDeviceResources`, and call `SetColor` in `DrawItem`.

---

**B10. `PasteShortcutFromClipboard` runs up to 256 `filesystem::exists` probes + `IPersistFile::Save` per source on the UI thread**
`RedSalamander/FolderView.FileOps.cpp:854`
Severity: **robustness** (UI-thread block) — CONFIRMED.
Mechanism: Invoked synchronously from the `WM_COMMAND` handler (`RedSalamander.cpp:9828`). Each source runs `GenerateShortcutPath`'s 256-iteration `exists()` scan (`:312`) plus a synchronous `IPersistFile::Save` (`:334`) and `CoCreateInstance(CLSID_ShellLink)`. (The 256 probes are worst-case; the dominant guaranteed cost is the synchronous `Save` per source.)
Failure: On a slow/offline/network destination the message loop freezes for seconds.
Fix: Move shortcut creation onto a worker / the async file-op pipeline, and cap the 256-attempt uniqueness scan with a faster strategy.

---

## 3. Cross-cutting themes

1. **Drop coordinate is never consulted.** D2/D3/B5 and the self-drop family all stem from `PerformDrop` having no `point` parameter and no hit-test. One fix (plumb `point`, hit-test, resolve folder-item destination, reject self/descendant) closes the largest cluster of wrong-target/data-loss findings.
2. **No self-drop / same-parent / descendant guard in the drop channel.** D1, D3, and B6's neighbors all rely on the FS plugin to do the right thing — and the plugin has no same-path early-out for directories. The guard belongs in `PerformDrop`, mirroring the existing file-ops bridge `isSameOrUnderFolder`.
3. **Trusting untrusted OLE/clipboard buffers.** S1/S2/S3 are the same anti-pattern: parse an HGLOBAL from another process with no `GlobalSize` bound. The internal-drop branch already demonstrates the correct `bytesAvailable`-bounded pattern; the CF_HDROP fallback, the `reserve`, and the preferred-drop-effect read each omit it. Standardize on `DragQueryFileW` and a `GlobalSize`-checked read helper.
4. **Async/UI-thread boundary violations.** R1 (worker reads UI-owned `_itemsFolder`), B1 (blocking shell I/O on the worker that the UI joins), B4 (untrusted count → thread terminate), B10 (synchronous shell I/O on the UI thread), and D7/D8 (acting on in-flight `_currentFolder` / focus migrated by an in-loop refresh) all share a root discipline gap: **operate on immutable snapshots taken at gesture time, and never block or join the UI thread on uncancellable I/O.**
5. **Stale source/target identity.** D5, D6, D7, D8 all let the *acted-on* set diverge from the *user-intended* set between gesture and execution. The cure is to capture the target identity (by stable name/id) at the moment of the gesture and re-validate at execution, refusing destructive ops when the target no longer resolves.
6. **MOVE reported before it happened.** D4 and B6 both report a MOVE the system has not verified — D4 to an external OLE source (data-loss), B6 as a convention default. Never report MOVE as performed until a copy is verified.

---

## 4. Coverage & residual risk

Static analysis confirmed the code mechanisms above; the following require **interactive / runtime** testing to characterize fully and were not exercised:

- **Live OLE drag-drop from a second app.** D4 (Explorer deleting the source on a reported MOVE) is OS/source-app behavior outside this repo — needs a real Explorer→virtual-FS MOVE with an induced copy failure (full disk / revoked permission / disconnected cloud) to observe actual loss. S1's crash-vs-bogus-path outcome depends on real adjacent-heap contents and a crafted/foreign data object.
- **Self-drop and descendant-move on real FS plugins.** D1's directory-destruction chain and D3's pre-existing-descendant path were traced through the FS plugin but should be reproduced on local NTFS, a 7z archive, and a cloud FS to confirm the copy+delete fallback and the absence of a prompt.
- **Teardown / navigation races.** R1 (the `_itemsFolder` torn read) and B1 (UI join blocking on `GetImage`) are timing-dependent — need a slow/disconnected SMB share or a deliberately-slow third-party thumbnail handler, with rapid A→B navigation and pane close mid-extraction, ideally under TSan/ASan or a heap-validation build.
- **Context-menu reentrancy (D8).** Requires an external process renaming/deleting the right-clicked item inside the ~200 ms debounce + menu-open window while enumeration completes in the nested loop. Scriptable but inherently racy; run repeatedly under stress.
- **Device-lost recovery (B7).** Needs a forced TDR (`dxcap`/driver reset) or a real dGPU↔iGPU switch / RDP transition while presenting, to confirm the pane never recovers.
- **Untrusted-codec exposure (B3) and the unbounded `reserve` (B4).** Need fuzzed/malicious images in a browsed folder and a stub third-party `IFileSystem` returning a hostile `GetCount`, respectively — neither is reachable via the in-tree plugins alone.
- **Cross-app paste (S3, B5, B10).** Need clipboards populated by other apps (undersized `CFSTR_PREFERREDDROPEFFECT`; CF_HDROP onto a virtual-FS pane; shortcut paste into a slow share) to confirm the OOB read, the source-reader misrouting, and the UI freeze respectively.

Highest residual risk if shipped without fixes: **D1 (directory self-drop destruction)** and **D2 (drop-point-ignored wrong-target MOVE)** — both reachable by ordinary single gestures with no special timing — followed by **D4** and the **S1/S2** OLE-parse defects, which an adversarial or merely buggy peer application can trigger.
