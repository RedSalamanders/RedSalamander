> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# NavigationView - verified findings (compact)

- **[contract/CONFIRMED]** `RedSalamander/NavigationView.Menus.cpp:2520` - Sibling-folder dropdown command IDs are unbounded and overflow the documented 600-699 range into the history ID range (700-799)
- **[contract/PLAUSIBLE]** `RedSalamander/NavigationView.Menus.cpp:2520` - Sibling dropdown assigns unbounded menu command IDs (ID_SIBLING_BASE + i) that overrun the reserved 600-699 range into the history range
- **[contract/PLAUSIBLE]** `RedSalamander/NavigationView.FullPathPopup.cpp:459` - Sibling-folder menu command IDs overflow the 600-699 range and collide with history IDs (700-799)
- **[dos-oom/CONFIRMED]** `RedSalamander/NavigationView.Edit.cpp:2320` - Address-bar commit blocks the UI thread on dead/offline UNC and network paths (synchronous GetFileAttributesW)
- **[dos-oom/CONFIRMED]** `RedSalamander/NavigationView.Edit.cpp:2320` - Address-bar submit blocks the UI thread on GetFileAttributesW for typed UNC / offline paths
- **[dos-oom/CONFIRMED]** `RedSalamander/NavigationView.Edit.cpp:2320` - Path commit validation does synchronous GetAttributes on the UI thread — UNC/offline path freezes the app
- **[dos-oom/CONFIRMED]** `RedSalamander/NavigationView.Edit.cpp:1382` - Synchronous plugin DLL load + Initialize on the UI thread during keystrokes when instance context changes
- **[dos-oom/CONFIRMED]** `RedSalamander/NavigationView.Edit.cpp:2320` - Typed/pasted address validation does synchronous GetFileAttributesW on the UI thread, freezing on slow/offline UNC paths
- **[dos-oom/CONFIRMED]** `Plugins/FileSystem/FileSystem.Menu.cpp:264` - Drive/menu dropdown enumerates all logical drives with synchronous GetVolumeInformation/GetDiskFreeSpaceEx/GetDriveType on the UI thread
- **[dos-oom/CONFIRMED]** `RedSalamander/NavigationView.Interaction.cpp:228` - Drive/menu button click synchronously enumerates all drive volumes on the UI thread (GetVolumeInformationW/GetDiskFreeSpaceExW) -> multi-second freeze on disconnected network/offline drives
- **[dos-oom/CONFIRMED]** `RedSalamander/NavigationView.Rendering.cpp:519` - UpdateMenuIconBitmap performs live shell filesystem I/O (SHGetFileInfo) on the UI thread for UNC/offline paths
- **[dos-oom/CONFIRMED]** `RedSalamander/NavigationView.Edit.cpp:2320` - Synchronous GetAttributes on path-edit submit freezes the UI thread for UNC / offline / network paths
- **[dos-oom/CONFIRMED]** `RedSalamander/NavigationView.Edit.cpp:2320` - Address-bar commit blocks the UI thread on synchronous filesystem I/O (UNC/offline freeze)
- **[dos-oom/CONFIRMED]** `RedSalamander/NavigationView.Menus.cpp:2514` - Siblings dropdown builds an unbounded menu from a directory listing (large-folder hang/bloat)
- **[dos-oom/PLAUSIBLE]** `RedSalamander/NavigationView.Breadcrumb.cpp:274` - Breadcrumb layout is O(segments^2) with one DWrite layout per segment and no cap on segment count
- **[dos-oom/PLAUSIBLE]** `RedSalamander/NavigationView.Breadcrumb.cpp:50` - Breadcrumb Present passes an unvalidated dirty rect that can be empty/negative-width, mapping to a device-discard thrash loop
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/NavigationLocation.h:299` - Unconditional %VAR% environment-variable expansion of typed file-plugin paths makes literal folders named with %TOKEN% unreachable / mis-targeted
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/NavigationView.Edit.cpp:776` - Escape that dismisses the suggestion popup does not cancel the in-flight request, so stale suggestions reappear
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/NavigationView.Edit.cpp:364` - OnEditSuggestResults reopens the popup without re-checking that the edit field text still matches
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/NavigationView.FullPathPopup.cpp:1185` - Popup edit validates the raw typed string but navigates the env-expanded path, rejecting valid %VAR% paths
- **[robustness/CONFIRMED]** `RedSalamander/NavigationView.FullPathPopup.cpp:298` - WM_ACTIVATE handler can destroy the popup while a child dropdown / focus transfer is in flight, leaving D2D target bound to a dead HWND
- **[robustness/PLAUSIBLE]** `RedSalamander/NavigationView.Edit.cpp:1429` - Per-keystroke synchronous LoadLibraryExW + IFileSystemInitialize::Initialize on the UI thread when typing a plugin instance context
- **[robustness/PLAUSIBLE]** `RedSalamander/NavigationView.Menus.cpp:1376` - Change Drive submenu assigns a live command id to command-only items but registers no action, producing dead no-op menu entries
- **[robustness/PLAUSIBLE]** `RedSalamander/NavigationView.Menus.cpp:1984` - History dropdown snapshot can navigate to a stale entry after async history mutation during the modal menu
- **[robustness/PLAUSIBLE]** `RedSalamander/NavigationView.Rendering.cpp:779` - RenderDiskInfoSection draws the history chevron into the History section's rect while presenting only the DiskInfo dirty rect
- **[robustness/PLAUSIBLE]** `RedSalamander/NavigationView.FullPathPopup.cpp:878` - Full-path popup segment layout is unbounded in segment count and line count for very long paths
- **[robustness/PLAUSIBLE]** `RedSalamander/NavigationView.Menus.cpp:2520` - Siblings dropdown generates unbounded menu command ids that overflow into history/command id ranges
- **[wrong-target/CONFIRMED]** `RedSalamander/NavigationView.Edit.cpp:2280` - Drive-relative input (C:foo) is accepted and navigated against the process current directory
- **[wrong-target/PLAUSIBLE]** `RedSalamander/NavigationView.Edit.cpp:861` - Unconditional %VAR% environment-variable expansion makes literal percent-named folders unreachable and can silently retarget navigation
- **[wrong-target/PLAUSIBLE]** `RedSalamander/NavigationView.cpp:709` - Posted plugin navigate request is replayed with no re-validation of the active plugin (wrong-target after plugin swap)
