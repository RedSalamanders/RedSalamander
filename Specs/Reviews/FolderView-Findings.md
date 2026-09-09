> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# FolderView - verified findings (compact)

- **[crash/CONFIRMED]** `RedSalamander/FolderView.DragDrop.cpp:532` - CF_HDROP from a foreign app is parsed with no buffer bounds checks (OOB read / crash)
- **[crash/CONFIRMED]** `RedSalamander/FolderView.DragDrop.cpp:460` - Attacker-controlled pathCount drives vector::reserve inside a noexcept function -> std::terminate
- **[data-loss/CONFIRMED]** `RedSalamander/FolderView.DragDrop.cpp:550` - Drop has no self-drop / same-source-folder guard; a MOVE back into the source folder routes a move-onto-self into the FS
- **[data-loss/PLAUSIBLE]** `RedSalamander/FolderView.DragDrop.cpp:549` - PerformDrop has no self-drop / drop-into-source-folder / recursive-into-descendant rejection
- **[data-loss/PLAUSIBLE]** `RedSalamander/FolderView.DragDrop.cpp:575` - Drop target returns DROPEFFECT_MOVE to an external source before the async move has run, deleting the source on a copy that may later fail
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/FolderView.DragDrop.cpp:564` - External local-file drop onto a virtual-FS (archive/cloud) pane is routed to the destination's filesystem as the SOURCE reader
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/FolderView.DragDrop.cpp:311` - Drop effect defaults to COPY regardless of volume, deviating from same-drive MOVE convention
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/RedSalamander/FolderView.Icons.cpp:1004` - Teardown/navigation hangs the UI thread on a blocking IShellItemImageFactory::GetImage for offline/network thumbnails
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/FolderView.DragDrop.cpp:572` - Drop always targets _currentFolder and ignores the drop point, so a no-modifier intra-app drag silently COPIES a selection into its own source folder instead of moving onto a hovered subfolder
- **[race/CONFIRMED]** `RedSalamander/FolderView.Icons.cpp:539` - Worker thread reads non-atomic _itemsFolder (std::filesystem::path) while UI thread move-assigns/clears it — data race / torn read of freed buffer
- **[race/CONFIRMED]** `RedSalamander/FolderView.Icons.cpp:911` - ProcessThumbnailLoadQueue reads non-atomic _itemsFolder on worker thread (same race as icon queue, second site)
- **[robustness/CONFIRMED]** `RedSalamander/FolderView.Rendering.cpp:1876` - Device-lost (D2DERR_RECREATE_TARGET / DXGI_ERROR_DEVICE_REMOVED) never recreates the D3D/D2D device; view stays permanently blank
- **[robustness/CONFIRMED]** `RedSalamander/FolderView.Rendering.cpp:2394` - Hover/background brush created with raw CreateSolidColorBrush per frame is fine, but per-item selected-text brush is recreated every DrawItem call in the hot loop
- **[robustness/CONFIRMED]** `RedSalamander/FolderView.FileOps.cpp:267` - ReadPreferredDropEffectClipboard dereferences a DWORD from a clipboard HGLOBAL without checking its size (OOB read)
- **[robustness/CONFIRMED]** `RedSalamander/FolderView.FileOps.cpp:854` - PasteShortcutFromClipboard performs up to 256 filesystem::exists probes plus IPersistFile::Save per source on the UI thread
- **[robustness/PLAUSIBLE]** `RedSalamander/FolderView.Enumeration.cpp:393` - Unbounded files.reserve()/directories.reserve() driven by untrusted virtual-FS GetCount → bad_alloc → std::terminate (DoS)
- **[robustness/PLAUSIBLE]** `RedSalamander/RedSalamander/FolderView.Icons.cpp:145` - WIC thumbnail decode runs an untrusted-image parser in-process on the worker without sandboxing or size sanity limits on source dimensions
- **[robustness/PLAUSIBLE]** `RedSalamander/FolderView.Rendering.cpp:2176` - DrawItem dereferences _items[_hoveredIndex] guarded only by sentinel check, not bounds — unlike every other _hoveredIndex use site
- **[security/CONFIRMED]** `RedSalamander/FolderView.DragDrop.cpp:532` - CF_HDROP path-list parse trusts attacker-supplied DROPFILES offset and walks with no GlobalSize bound (OOB read)
- **[security/CONFIRMED]** `RedSalamander/FolderView.DragDrop.cpp:532` - PerformDrop walks a CF_HDROP path list from another app with no GlobalSize/terminator bounds check (OOB read)
- **[wrong-target/CONFIRMED]** `RedSalamander/FolderView.DragDrop.cpp:572` - Drop ignores the drop point: files always land in the current folder, never the folder item under the cursor
- **[wrong-target/CONFIRMED]** `RedSalamander/FolderView.Interaction.cpp:477` - Click-drag of an unselected item moves the previously-selected items (wrong-target data movement)
- **[wrong-target/CONFIRMED]** `RedSalamander/FolderView.Interaction.cpp:483` - Drag started on empty background drags the previously-focused item (wrong-target file movement)
- **[wrong-target/CONFIRMED]** `RedSalamander/RedSalamander/FolderView.Icons.cpp:794` - Thumbnails are extracted via SHCreateItemFromParsingName/WIC for virtual-FS (archive/cloud/curl) items using a fabricated local-looking path
- **[wrong-target/PLAUSIBLE]** `RedSalamander/FolderView.Menus.cpp:372` - Context-menu command captures the right-clicked item but executes against the selection re-read after an async refresh (stale/wrong target)
- **[wrong-target/PLAUSIBLE]** `RedSalamander/FolderView.DragDrop.cpp:572` - Drop/paste destination uses requested _currentFolder while displayed items belong to _displayedFolder (wrong-target during async navigation)
- **[wrong-target/PLAUSIBLE]** `RedSalamander/FolderView.Menus.cpp:367` - Context menu nested message loop lets an async refresh migrate focus to a different file, so the posted Delete/Move targets the wrong item
