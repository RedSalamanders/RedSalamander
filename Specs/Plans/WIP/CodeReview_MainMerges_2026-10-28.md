# I50 Review of main merges 2026-10-18..2026-10-28

Status: ACTIVE. Review record and remediation routing. No product code has been changed by this plan.
Owner: next engineering owner chosen by the product owner. Reviewed at `8ea956a7e` (main) against the
pre-window baseline `0f6cd8ddb`. Date: 2026-10-28.

Drift check (run first):

```powershell
git diff --stat 8ea956a7e..HEAD -- RedSalamander Plugins Common Tools .github build.ps1 Directory.Build.props Directory.Build.targets
```

Any in-scope file that changed after `8ea956a7e` must be re-read against the cited excerpt before its
slice is executed. Line numbers below are on `8ea956a7e`.

## Scope

| Merge | Content | Size |
|---|---|---|
| PR #142 (`5c9292498`, 2026-10-26) | Source snapshot `397e15adf` of DualTail up to `582d644be` (about 2026-09-19..10-26: File Operations UX/reliability I35, Issues pane, decision mirror, leftover cleanup, ViewerSpace workspace/analysis, RedMonitor extraction, MSVC 14.52 + Release LTCG, DxUi API 4, tooling) plus 10 follow-ups (hosted MSVC, capture WinRT, ViewerSpace scan queue, ViewerWeb profile/lifetime) | 25,590 files; production code about 304 files, +65k/-12.5k lines (renames detected, CR ignored); tests about +110k; `Specs/TestRuns` 19,500 evidence files |
| PR #144 (`e101eac66`) | Agent/skill audit record (docs) | 6 files |
| PR #143 (`fbcc9fc0d`) | Hosted capture admission, S3 fixture latency (test-only) | 8 files |
| PR #145 (`8ea956a7e`) | PowerShell formatter tooling, agent guidance/skill trims, spec IA checks | 45 files |

Reviewed: all production C++ (RedSalamander, Common, Plugins, RedConfigure), build/CI/installer,
Tools modules and entry points, test-harness code that gates qualification. Not reviewed in depth:
test-case bodies (about 110k lines), evidence under `Specs/TestRuns`, plan prose under `Specs/Plans`,
the external pinned DxUi package, translation quality.

Method: 12 area reviewers (Sonnet, read-only) with a shared brief, then lead verification. Every
finding below carries one verification label:

- **CONFIRMED**: the lead re-read the cited code and the trigger chain.
- **AGENT**: the reviewer confirmed it by reading; the lead did not re-trace it.
- **PLAUSIBLE**: the mechanism is confirmed but the user-visible effect needs a runtime check.

"New" means introduced in the window; "pre-existing" means present at `0f6cd8ddb` (still a production
defect, listed because the review covered the surrounding code).

## Verdict

`main` at `8ea956a7e` is **not production-ready yet**. No data-loss path was found in the Local
provider's delete, overwrite, Move or Recycle code, in the leftover-cleanup Delete, or in the File
Operations result reducer and conflict broker. What blocks a production claim:

1. **No green qualification at the merged identity** (X1). The scheduled Nightly Native Suite (all
   four profiles), weekly ASan, Nightly Flake Convergence and Ghostty status workflows are red on
   `main`. The merged I47 (RedMonitor) and I48 (MSVC 14.52) plans still record "final six builds,
   three Fresh Full profiles, measurements and closeout remain pending". The last recorded local
   Debug Fresh Full (2026-10-24) failed (Commands 951/2/6, FileOps 278/4/40, ViewerPE interactive).
   Most of the nightly red comes from the hosted runner (1024x768 display, Windows Terminal owns the
   foreground) and harness defects (H2, M28, M31, M32), not the product. Three product failures
   reproduce there and also failed at the baseline (M29, M30, M33). An in-flight branch,
   `ej/ci-runtime-repair-20261028` (plan `HostedCiRepair_2026-10-28.md`, not on `main`), already owns
   the hosted CI repair; route CI items there instead of duplicating it.
2. **Three High defects**: plugin host services that return success without doing anything when
   called off the UI thread (H1); MSIX packaging that always fails (H3); a repository that no longer
   checks out on default Windows git, which breaks the release and several scheduled workflows (H4).
3. **34 Medium defects**, several with direct user impact: ViewerSpace Delete of a folder,
   silent loss of saved popup/Issues settings, Hidden attribute lost on Overwrite, sticky MTP
   fallback that drops atomic creates, lost-wakeup hangs, a running Move shown as interrupted.

## High

| ID | Area | New? | Verified | Finding | Where |
|---|---|---|---|---|---|
| H1 | Host services | pre-existing | CONFIRMED (routing); PLAUSIBLE (MS Drive effect) | `IHostConnections::SetConnectionSecret`, `DeleteConnectionSecret` and `IHostViewers::OpenViewer` called off the UI thread do `SendMessageW` to the FolderWindow, whose `WndProc` has no case for `kHostSetConnectionSecret`/`kHostDeleteConnectionSecret`/`kHostOpenViewer`; `DefWindowProcW` returns 0, the sender returns it as `S_OK`. Microsoft Drive stores rotated/new OAuth refresh tokens this way from provider operations, so a token can silently never be saved or deleted (repeated sign-in). | `RedSalamander/FolderWindow.cpp:1404-1413` (only dispatcher), `RedSalamander/HostServices.cpp:907,934,1216` (senders), `:1332,1346,1419` (unreachable handlers), caller `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:7909,7936,7956`, reached from `MicrosoftDriveFileWriter::UploadSimple` (`:7516-7519`) on a copy worker |
| H2 | Test harness | new | CONFIRMED (code + CI log) | ViewerPE desktop-restoration identity requires the saved focus HWND to belong to the foreground window's thread. A focus child created by another thread of the same process (same pid, same GA_ROOT) fails `savedIdentityValid`, the no-write shortcut is also skipped, and the Viewer suite aborts. This is the ARM64 nightly failure (`savedFocusTid=1448`, `savedTid=7896`, same pid 3196, input unchanged). | `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:1021-1044` |
| H3 | Packaging | new | CONFIRMED (target condition; reviewer ran the target) | `ValidateMsvcToolsVersion` runs for every first-party project, including the MSIX `.wapproj`, which imports no VC props, so `VCToolsVersion` is empty and the build stops with "requires MSVC build tools 14.52.x; selected ." No MSIX can be produced (`build.ps1 -Msix`, `release.yml` package-msix). `RSVerifyRedMonitorClient` beside it already filters on `.vcxproj`. | `Directory.Build.targets:324-326` |
| H4 | Repository / CI | new | CONFIRMED | 607 tracked paths (up to 288 characters) under `Specs/Plans/Done/.../visual-gallery/` and `Specs/TestRuns` exceed MAX_PATH on a default Windows checkout; none did at the baseline. Workflows without `core.longpaths` fail at Checkout: Ghostty runtime status and Nightly Flake (logs), and by the same mechanism every `release.yml` job and `winget-release.yml` (not yet run). Contributors on stock Windows git cannot clone. | `.github/workflows/ghostty-runtime-status.yml:26-31`, `nightly-flake.yml:19-24`, `release.yml:42,75,174,219,404`, `winget-release.yml:77,83` |
| X1 | Qualification | process | CONFIRMED | See verdict item 1. Runs: Nightly Native 40660649657 (build passes; suite fails 17-18 cases on x64 and 59-66 on ARM64), ASan 40484804050 (ARM64 59 failures; x64 FileOps suite silent for 4 h 10 min until the 6 h job limit, cause not diagnosed), Flake 40691121033, Ghostty 40665991855. Earlier "successful" nightlies were 8-second skips; the only real baseline run (35498140593 at `0f6cd8ddb`) failed the same cluster. The "Manifest validation failed" and "Merge of test-triplet" log lines come from passing Pester tests, not failures. | `.github/workflows/*`, `Specs/Plans/WIP/Msvc1452ReleaseOptimization_2026-10-23.md`, `RedMonitor_Extraction_2026-10-21.md` |

## Medium

| ID | Area | New? | Verified | Finding | Where |
|---|---|---|---|---|---|
| M1 | Viewer actions | new | CONFIRMED (code); PLAUSIBLE (runtime) | ViewerSpace "Delete" (and recycle escalation) is admitted with `FILESYSTEM_FLAG_NONE`, so `deletion.recursive=false` and a marked non-empty folder fails with `ERROR_DIR_NOT_EMPTY` after the user confirmed. Pane and Find Files delete pass `FILESYSTEM_FLAG_RECURSIVE`. | `RedSalamander/FolderWindow.FileOperations.cpp:3722-3728`, `:2629` |
| M2 | FileOps engine | pre-existing | CONFIRMED (mechanism) | Lost-wakeup class: `PerItemTaskScheduler::NotifyWorkAvailable` and `Shutdown`, `NotifyQueueChanged` (used by `RequestCancel`), and the bridge pipeline `readerFinished`/`pipelineStop` stores notify without taking the waiter's mutex while the predicates read state written elsewhere. Waits are untimed, so Resume/Cancel can be lost and a job, queued task or `readerThread.join()` hangs. Same class as the earlier `SetPaused` bug; `WakePauseWaiters` already shows the fix. | `RedSalamander/FolderWindow.FileOperations.State.cpp:6479-6482`, `:6497-6525`, `:6899`, `:12660-12692`, `:18394-18395`, `:18490-18493`; `State.Runtime.cpp:2560-2567`; `State.Queue.cpp:1850`, `:2041` |
| M3 | Preferences | pre-existing (+1 key new) | CONFIRMED | `MaybeResetWorkingFileOperationsSettingsIfEmpty` hand-copies the non-default test and omits `popupFooterOnly`, `popupCompactDensity`, `popupThroughputExpanded` and the three `issuesPane*` keys. Turning a Preferences option back to default resets the whole `fileOperations` object and Apply saves it: popup density/disclosure and the Issues pane sort/columns are silently lost. | `RedSalamander/Preferences.Internal.cpp:1469-1491` vs `Common/Common/SettingsStore.cpp:5517-5535` |
| M4 | Local provider | pre-existing | CONFIRMED | Overwrite (staged) copies and preserved links lose Hidden: `ClearTemporaryAttributes` strips `HIDDEN|TEMPORARY` from the owned stage after the source attributes were applied. A hidden file copied over an existing one ends visible; a new-name copy keeps Hidden. | `Plugins/FileSystem/FileSystem.cpp:4408-4431`, `:4607`; `FileSystem.FileOps.cpp:4594-4635`, `:5719-5750` |
| M5 | Local provider | pre-existing | CONFIRMED | Stage and backup names are `<leaf>.rs_copy_tmp_<pid>_<tid>_<counter>` (+47 chars) and `<leaf>.rs_bak_{GUID}` (+46). Leaves of 209+ characters fail Overwrite/new-directory/link copy with an invalid-name error; the names are predictable. The spec forbids PID/TID/counter fallbacks; `Common::Paths::BuildUniqueSiblingName` already does this right. | `Plugins/FileSystem/FileSystem.FileOps.cpp:4184-4192`, `FileSystem.cpp:4389-4406`; `Specs/FileSystem/FileSystem_FileOperations.md:1074-1075` |
| M6 | Local provider | pre-existing | CONFIRMED | Link Preserve rebuilds the target from the advisory PrintName instead of the SubstituteName, and `Win32PathToNtPath` maps `\\?\C:\x` to `\??\UNC\?\C:\x`, so such links are copied dangling or pointing elsewhere. | `Plugins/FileSystem/FileSystem.FileOps.cpp:11280`, `:2398-2411` |
| M7 | Local provider | pre-existing | PLAUSIBLE | Direct and staged copy fail with `ERROR_FILE_INVALID` when the path-query size differs from the bound reader's size (file being appended, or changed while waiting for a transfer slot); Retry fails the same way. | `Plugins/FileSystem/FileSystem.FileOps.cpp:5473-5509`, `:5221-5224`, `:4429-4432` |
| M8 | MTP | new | CONFIRMED | The "create temp refused" memo is set before the direct fallback, includes `ERROR_ACCESS_DENIED`, is never reset, and `SupportsAtomicWriterCommit` keeps answering TRUE. One denied folder turns every later create on the device into a non-atomic direct upload for the rest of the session. | `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:268-273`, `:3524-3541`, `:4532-4543`, `:3309-3326` |
| M9 | Move breadcrumb | pre-existing | CONFIRMED | Every `FileOperationState` ctor loads all non-terminal breadcrumbs, and records carry no writer identity. Opening Compare Directories (own embedded FolderWindow) or a second process during a long Move shows a false "Move was interrupted" card; dismissing it deletes the live record, so a crash later produces no notice. | `RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp:567-571`, `:578-665`; `FileOperationMoveBreadcrumb.cpp:486-508`; `FileOperationMoveBreadcrumb.h:50-61` |
| M10 | Move breadcrumb | new | CONFIRMED | The interrupted-Move card header uses `sourceRootCount` (unique provider roots, usually 1) as the item count and a full path as the single name, so a 500-file interrupted Move reads "Move C:\...\first.txt". | `RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp:629-631`, `:650-652` |
| M11 | FileOps engine | pre-existing | CONFIRMED (early exits); AGENT (classification) | Inline Rename refusals before any mutation (file in use, vanished, access denied, case-only unsupported) return before `recordStatus` exists, so the task ends Indeterminate / `ERROR_IO_INCOMPLETE` ("outcome unknown") instead of a clean "not renamed". Batch Rename handles this case. | `RedSalamander/FolderWindow.FileOperations.State.cpp:13505-13548` |
| M12 | Popup keyboard | new | CONFIRMED | With no card control focused, Escape is bound to the sole actionable decision in the whole snapshot, even when that card is scrolled out of view, and submits its Escape action instead of hiding the popup. A user pressing Escape to close the popup can cancel a task they cannot see. Spec allows the sole-decision fallback but does not require visibility (Q4). | `RedSalamander/FolderWindow.FileOperations.Popup.cpp:8393-8460`, `:8479-8482`, `:13931-13938` |
| M13 | Popup a11y | pre-existing | CONFIRMED | A plain Running task's progress bar has an empty UIA Name and the failure surface prints "Copy: " with nothing after it: `StatusTextForTask` returns `{}` for Running and the fallback only covers a missing task. | `RedSalamander/FolderWindow.FileOperations.Popup.cpp:8711-8720`, `:1765-1810`, `:13807-13810` |
| M14 | Plugin buffers | new (leftover) + pre-existing (pane, sibling menu) | CONFIRMED | Three walks of plugin `IFilesInformation` buffers do not bound `FileNameSize` (and the sibling menu not even `NextEntryOffset`); the leftover walk also joins unvalidated child names with `/=`, so a rooted or `..\` name escapes the folder and is offered pre-checked for deletion. `Common::Plugins::PackedFileInfoCursor` exists for exactly this. | `RedSalamander/FolderWindow.LeftoverCleanup.cpp:683-709`; `FolderView.Enumeration.cpp:424-516`; `NavigationView.Menus.cpp:2394-2452` |
| M15 | Leftover cleanup | new | CONFIRMED (mechanism) | `~LeftoverCleanupDialogState` calls `StopScan(true)`, which loops `CancelSynchronousIo` + `Sleep(20)` on the UI thread with no bound. `CancelSynchronousIo` does not interrupt Curl/HTTP socket waits, so closing the window or the app after a subfolder walk on a stalled FTP/SFTP/S3/Drive location hangs until the provider times out. Non-final closes also sleep up to 2 x 200 ms. | `RedSalamander/FolderWindow.LeftoverCleanup.cpp:77-91`, `:126-129`, `:800-817`, `:1104-1137`; `FolderWindow.cpp:2286-2291` |
| M16 | Leftover cleanup | new | CONFIRMED | Unreadable folders in the subfolder walk are skipped silently; if the root fails the dialog shows "No leftovers ... found here" after it listed rows a moment before. | `RedSalamander/FolderWindow.LeftoverCleanup.cpp:684-692`, `:767-794`, `:544-585` |
| M17 | Leftover cleanup | new | PLAUSIBLE | The "row in use" filter is textual and window-local, a live scope without `pathIdentity` fails open, and pre-check ignores age, so another window's or process's live stage can be offered pre-checked. No user data is lost (a stage is a copy) but another transfer can fail. | `RedSalamander/FolderWindow.LeftoverCleanup.cpp:336-351`; `State.Queue.cpp:1590-1636` |
| M18 | FolderView | pre-existing | PLAUSIBLE | Cancel on a busy enumeration only bumps the generation; the provider call keeps the single enumeration worker, so the next navigation waits until the redirector times out. Spec says Cancel "aborts enumeration" (Q6). | `RedSalamander/FolderView.Enumeration.cpp:329`, `:928`; `FolderView.Interaction.cpp:593`; `DirectoryInfoCache.cpp:23-48` |
| M19 | ViewerSpace | new | CONFIRMED (code) | `_modeCombo` is never assigned, so `GetFocusControl() == _modeCombo` is true whenever no DxUi control has focus; Up/Down in the navigation row then silently switch Explorer/Analyst and persist the configuration. | `Plugins/ViewerSpace/ViewerSpace.NavigationHost.cpp:403`, `ViewerSpace.NavigationHost.h:121` |
| M20 | ViewerSpace | new | CONFIRMED | `WorkspaceOwnedBytes()` is counted inside `AnalysisOwnedBytes()` and again in `ViewerOwnedBytes()` (and in deferred-disposal charges); a process-global 8 MiB export charge inflates every viewer, so refresh/cache/export are refused near the budget. | `ViewerSpace.Analysis.cpp:235`, `ViewerSpace.Layout.cpp:247,271`, `ViewerSpace.Model.cpp:88`, `ViewerSpace.Workspace.cpp:1134-1146` |
| M21 | ViewerSpace | pre-existing | CONFIRMED | Up at a UNC share root is enabled and rescans `\\server\` (MSVC `root_path` of `\\nas\media` is `\\nas\`), ending in Error with the good model discarded; a trailing separator rescans the same folder. Spec says Up is disabled at a share root. | `Plugins/ViewerSpace/ViewerSpace.cpp:1411-1438`, `:13715-13724`, `:13852-13861` |
| M22 | ViewerSpace | pre-existing | CONFIRMED | Closing during a scan tests `done` immediately after `request_stop`; a cooperative worker is almost always still running, so it is detached and the unload gate is kept, quarantining plugin unload for the rest of the process. The "uncooperative provider" exception is the common close path. | `Plugins/ViewerSpace/ViewerSpace.cpp:8436-8503`, `:2238` |
| M23 | FileOps queue | pre-existing | PLAUSIBLE | The same-host overlap advisory runs the pairwise O(N x M) scope predicate while holding `_queueMutex`, blocking workers and the UI (Start now, queue moves) for large selections. The gate already has an indexed path. | `RedSalamander/FolderWindow.FileOperations.State.Queue.cpp:1168-1219`, `:1432-1526` |
| M24 | Host services | pre-existing | CONFIRMED | Eight synchronous host messages still cast `lParam` to a raw pointer (any same-desktop process can crash the UI thread with one `SendMessage`); only two were moved to a token registry whose tokens are a sequential counter. Contradicts `Plugins_PluginAPI.md:188` and AGENTS.md; `docs/dev/PluginHostModel.md:107-108` documents the old pattern. | `RedSalamander/HostServices.cpp:116-218`, `:1295-1421` |
| M25 | Tooling | new | CONFIRMED | `format-all.ps1` formats every `*.h` outside build folders, including the hash-locked LF-only `ThirdParty/RedMonitor/include/RedMonitor/Client.h`; `.clang-format` forces CRLF and `ThirdParty/RedMonitor/.gitattributes` is `* -text`, so the lock hash breaks and every first-party project fails to build. It also uses the newest local clang-format, not the pinned 22.1.3. | `format-all.ps1:48-137`, `.clang-format:7`, `Dependencies/RedMonitor.lock.json` |
| M26 | Tooling | new | CONFIRMED | An exception in the end-of-run evidence archive (leftover junction, locked file, unstable inventory) is uncaught after the evidence run was already completed as passed: no summary, no verdict, exit 1 instead of the documented 5. | `Tools/Run-AllTests.ps1:2291`, `:2306-2307`, `:2588-2611` |
| M27 | Tooling | new | CONFIRMED | `-NonInteractive` / `Test-Changes.ps1 -Suite <family>` report PASSED and exit 0 for legacy suites although desktop entries were dropped; help promises NOT_EVALUATED / exit 2. | `Tools/Run-AllTests.ps1:1491-1499`, `:1746`, `:2356-2371`; `Test-Changes.ps1:38-62` |
| M28 | Test harness | new | PLAUSIBLE | x64 nightly restoration of a foreign foreground window is one unsettled `SetForegroundWindow` attempt, judged 1.4 ms later with no diagnostics, and its failure aborts the Viewer suite. | `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:1088-1112` |
| M29 | Search index | pre-existing | AGENT (reproduced by the nightly and the baseline run) | When the USN journal is unavailable, the rebuild calls `BuildIndex`, a whole-volume MFT enumeration, instead of the tracked-root traversal; `local_index_core_journal_replay_rename_delete_create` fails. | `Common/LocalSearchIndexCore.cpp:3375`, `:3479` |
| M30 | Search service | pre-existing | AGENT (reproduced) | Multi-client service access fails with `0x800700E7` (pipe busy) under the 750 ms connect limit; `search_service_multi_client_and_rebuild_control` fails on hosted runners. | `Common/SearchServiceBroker.cpp` |
| M31 | CI | pre-existing | PLAUSIBLE | The Nightly Flake `-SkipBuild` consumer can never use the receipt: `Test-RSBuildReceiptForArtifactUse` needs the pinned DxUi source tree, which the consumer job never restores, and returns a bare `$false` that does not name the failed predicate (100 of 100 runs failed since 2026-07-22). | `Tools/Modules/Build/BuildEvidence.psm1:993-1018`, `Tools/Run-AllTests.ps1:1432-1441`, `.github/workflows/nightly-flake.yml` |
| M32 | CI | pre-existing | AGENT | When the test-root step did not run, the Nightly Flake upload uses `path: \` and scans the drive root (`EPERM ... System Volume Information` masks the real error); on success it uploads the whole sandbox, including VHDX evidence, without the reusable workflow's exclusions. | `.github/workflows/nightly-flake.yml:84-90` |
| M33 | FileOps throughput | pre-existing | AGENT (reproduced) | `FileOps_SelectionShapeLocal` fails on 4-core runners: selected-file copy is capped at 3 transfers against 4 for folders because the per-item pool sizing counts the starvation reserve. | `RedSalamander/FolderWindow.FileOperations.State.cpp:6630-6637` |
| M34 | CI / build | new | PLAUSIBLE | The inline CI `vcpkg install` passes neither the overlay triplets nor the ARM64 `aws-c-common` overlay port, and the cache key omits `Build/vcpkg-ports`; a cold ARM64 cache would compile stock `aws-c-common` with the C1001 the overlay avoids. Every observed run restored from cache, so this path is untested. `build.ps1` comments claim CI uses the governed installer. | `.github/workflows/build-reusable.yml:381-388`, `Build/vcpkg-ports/aws-c-common/README.md` |

## Low

| ID | Area | New? | Verified | Finding | Where |
|---|---|---|---|---|---|
| L1 | ViewerSpace | new | CONFIRMED | Retained refresh has no visible Cancel button, but a hidden Cancel hit target in the header discards it on click. | `ViewerSpace.cpp:4072-4088` vs `:5622-5626` |
| L2 | ViewerSpace | new | PLAUSIBLE | Clicking the treemap never takes keyboard focus, so Backspace/arrows/Enter go to the filter or grid. | `ViewerSpace.cpp:5608-5652` |
| L3 | ViewerSpace | pre-existing | CONFIRMED | Arrow-key selection walks culled (invisible) draw items; Up/Down equal Left/Right. | `ViewerSpace.cpp:5442-5460` |
| L4 | ViewerSpace | new | PLAUSIBLE | Alt+Left/Right/Up/Insert arrive as `WM_SYSKEYDOWN` and are dropped when the canvas has focus. | `ViewerSpace.cpp:3318-3335` |
| L5 | ViewerSpace | new | CONFIRMED | `IDS_VIEWERSPACE_HEADER_PROBLEMS` is English in all four satellites ("Hotovo · 3 unreadable"); the translated `..._DONE_WITH_UNREADABLE_FORMAT` is unused. | `ViewerSpace.cpp:7366-7375`, `ViewerSpaceResources.rc:85,252` |
| L6 | ViewerSpace | new | CONFIRMED | `_workspaceScanStartedSeconds` is not reset by `StartScan`, so a restarted scan's items/s uses the previous start time. | `ViewerSpace.cpp:7417-7433`, `:8095-8109` |
| L7 | ViewerSpace | pre-existing | AGENT | A canceled scan stuck in a provider call keeps the volume permit; the next scan shows "Queued..." without explanation. | `ViewerSpace.cpp:1684-1729`, `:9196-9220` |
| L8 | ViewerSpace | pre-existing (heavier now) | AGENT | Natural sort, summary encoding and per-child record building run under `frontierMutex`. | `ViewerSpace.cpp:10220-10422` |
| L9 | ViewerSpace | new | CONFIRMED | File-count weighting evicts/sorts by bytes while area uses file count. | `ViewerSpace.Layout.cpp:356-357,470,536` |
| L10 | ViewerSpace | new | CONFIRMED | Breadcrumb folder menu takes the first 512 children in arena order, then sorts. | `ViewerSpace.cpp:3711-3735` |
| L11 | ViewerSpace | new | CONFIRMED | `_exportProjectionRequested` is not cleared on the WAIT/LIMIT early returns, so the alert re-shows each tick and the timer never idles. | `ViewerSpace.Workspace.cpp:78-112`, `:259-263` |
| L12 | ViewerSpace | new | AGENT | Deferred refresh adoption retry drops the "view folder selected" resolution. | `ViewerSpace.Model.cpp:526-530`, `:615` |
| L13 | ViewerSpace | new | AGENT | One unpaired-surrogate name fails the whole export with a generic I/O error. | `ViewerSpace.Workspace.cpp:296-303` |
| L14 | ViewerSpace | new | AGENT | Escape inside the analysis panel never reaches its own Escape handler (dead `panelClose` path or spec mismatch, Q7). | `ViewerSpace.AnalysisHost.cpp:195-202`, `:722-733` |
| L15 | FileOps engine | new | AGENT | 64-slot write-only consent-receipt table turns the 65th per-item consent into `ERROR_NOT_ENOUGH_MEMORY`. | `State.cpp:6145-6182`, `:6235-6246` |
| L16 | FileOps engine | pre-existing | AGENT | `Task::_stopToken` assigned on the worker while UI threads read it (formal data race). | `State.cpp:11824` |
| L17 | FileOps runtime | pre-existing | CONFIRMED | Interrupted-Move notices are evicted by the 24-summary cap although Clear keeps them (they return next launch). | `State.Diagnostics.cpp:2327-2335` |
| L18 | FileOps runtime | pre-existing | AGENT | Rejected/corrupt breadcrumbs are never logged (dead warning branch); failed `Advance(Executing)` logs at Debug. | `State.Runtime.cpp:582-592`, `State.cpp:12006-12018` |
| L19 | FileOps queue | pre-existing | AGENT | Indexed overlap scan sorts keys by raw string order (`DIR-OLD` sorts between `DIR` and `DIR\X`), so an ancestor can be popped early. Same weakness in Local `DeleteItems` (not reachable: host rejects nested selections). | `State.Queue.cpp:1043-1083`; `FileSystem.FileOps.cpp:13608-13647` |
| L20 | FileOps completion | pre-existing | CONFIRMED | Directory-cache notifications use the live pane's file system instead of the task's, leaving stale entries after navigation. | `FolderWindow.FileOperations.cpp:5255-5330` |
| L21 | Localization | pre-existing | AGENT | Hard-coded English plural `s` feeds `{1:s}` in cs/sk/ja confirmation text ("3 položkas"). | `FileOperationConfirmation.cpp:63-87`, `State.Runtime.cpp:1312-1337` |
| L22 | Diagnostics | pre-existing | AGENT | Issues export leaves a truncated file on write failure; `OpenDiagnosticsLogForTask` scans the Logs folder under `_mutex`. | `State.Diagnostics.cpp:645-649`, `:803-891` |
| L23 | Popup | mixed | AGENT | User-visible strings stitched with hard-coded separators (`" · "`, `": "`, `"{0}. {1}"`). | `Popup.cpp:11871`, `:13794-13810`; `PromptText.cpp:59-122` |
| L24 | Popup | pre-existing | AGENT | Arrow keys stop at a disabled control instead of skipping it. | `Popup.cpp:7646-7668` |
| L25 | Popup | new | AGENT | Focus on a Completed-group header control jumps to Options when the header scrolls out. | `Popup.cpp:9700-9703`, `:8839-8847` |
| L26 | Popup | unknown | PLAUSIBLE | Speed-limit prompt destroys `_dxHost` before `_hWnd`; WM_NCDESTROY then dispatches into the destroyed host at shutdown. | `Popup.cpp:3286-3287`, `:2582-2613` |
| L27 | Popup | pre-existing | PLAUSIBLE | "Cancel all" consent defaults to OK with no held-Enter guard (the failure-surface variant has one). | `Popup.cpp:12810-12859` |
| L28 | Popup perf | pre-existing | AGENT | Hidden popup rebuilds a full snapshot and calls `ITaskbarList3` at 10 Hz without change detection. | `Popup.cpp:14305-14309`, `:12496-12539` |
| L29 | Popup perf | new | PLAUSIBLE | Middle truncation re-runs a DirectWrite binary search for every path row every frame. | `Popup.cpp:3399-3464` |
| L30 | Overlay | new | PLAUSIBLE | `AlertOverlayWindow::WndProc` forwards only four messages to its DxUi host, against the DxUi "forward every message" contract. | `Ui/AlertOverlayWindow.cpp:1223-1310` |
| L31 | Specs | new | CONFIRMED | `Core_StartupBootstrap.md:105-117` still specifies the removed RedSalamanderMonitor auto-launch. | `Specs/Core/Core_StartupBootstrap.md` |
| L32 | MTP | new | AGENT | Create temps (`.rs-mtp-create-*.tmp`) have no recovery journal; a crash leaves a full-size hidden object on the phone. | `FileSystemMtp.Core.cpp:307-367` |
| L33 | MTP | new | CONFIRMED | `FileSystemMtp.Tests.FakeBackend.cpp` (~1,300 lines) has no `ENABLE_TESTS` guard and ships in the Release DLL; every sibling fake is guarded. | `Plugins/FileSystemMtp/FileSystemMtp.Tests.FakeBackend.cpp:1`, `FileSystemMtp.vcxproj:146` |
| L34 | ViewerVLC | pre-existing | AGENT | Rapid selections during the initial libVLC load spawn parallel loaders. | `ViewerVLC.cpp:5333-5400` |
| L35 | ViewerWeb | new (test builds) | CONFIRMED | Test-profile admission failure is silent (no alert, no log). | `ViewerWeb.cpp:4202-4214` |
| L36 | Local provider | new | AGENT | `\\?\unc\` (lower case) is UNC for routing but not for the prefix helpers. | `FileSystem.cpp:961-964`, `FileSystem.Path.cpp:29-32,266-369` |
| L37 | Local provider | pre-existing | PLAUSIBLE | An abandoned `Win32FileWriter` deletes by path after closing and swallows failure. | `FileSystem.cpp:4931-4943` |
| L38 | Diagnostics | new | CONFIRMED | ETW messages over 16,000 code units are now dropped (RedMonitor `ERROR_BUFFER_OVERFLOW`) instead of truncated. | `Common/Helpers.h:1715-1721`, `ThirdParty/RedMonitor/include/RedMonitor/Client.h:106-110` |
| L39 | Search | new | AGENT | Dead `state == Ready` branch and always-true `storedVolumeReady` probe after the new early return. | `Common/LocalSearchIndexCore.cpp:4012-4036` |
| L40 | Tooling | new | CONFIRMED | Changelog fold drops a distinct fragment whose text equals any Unreleased entry, then deletes its file. | `Tools/Modules/Documentation/Changelog.psm1:94-101` |
| L41 | Tooling | new | AGENT | `New-RSTestSandboxLink` validates a relative `-Path` against the process directory but creates it against the PowerShell location. | `Tools/Modules/Testing/TestSandbox.psm1:341-396` |
| L42 | Build | new | AGENT | `RSVerifyRedMonitorClient` spawns `pwsh` (~1.1 s) for each of 114 projects on every build. | `Directory.Build.targets:327-329` |
| L43 | Repo hygiene | new | CONFIRMED | Public repo now tracks 19,500 evidence files (pack 502 MB); the private NAS hostname appears in 3 Done plans and 27 evidence files; `advisor-plans/` (agent executor plans) and `.codex/config.toml` (`workspace-write`, `network_access = true`) were published. | `Specs/TestRuns`, `Specs/Plans/Done`, `advisor-plans/`, `.codex/config.toml` |
| L44 | Repo hygiene | pre-existing | CONFIRMED | Two different files share the name `FolderWindow.FileOperations.Tests.OwnedShutdown.cpp` (compiled hooks in `RedSalamander/`, included fixture in `SelfTest/FileOperations/`); `vcpkg.json` homepage typo `Redsalamanderse`. | `RedSalamander/RedSalamander.vcxproj:413,530`; `vcpkg.json:4` |
| L45 | Build | new | CONFIRMED (reading) | `Verify-RedMonitorClient.ps1` rejects any reparse point up to the drive root, so a checkout below a junction or OneDrive placeholder fails every build; it needs `pwsh` on PATH even for IDE builds (with L42). | `Tools/Verify-RedMonitorClient.ps1`, `Directory.Build.targets:327-329` |
| L46 | Monitor removal | new | AGENT | Dangling RedSalamanderMonitor remnants (none breaks the build): dead `g_redSalamanderMonitor*` and `kColorTextView*` constants, stale tool/test lists, ETW scripts that still say "launch RedSalamanderMonitor", docs. `.sln`, projects, installers and workflows are clean. | `Common/Helpers.h:1392-1393`, `Common/WindowMessages.h:1051-1054`, `init-etw-trace.ps1`, `start-etw-trace.ps1`, `stop-etw-trace.ps1`, `docs/SettingsFile.md`, `docs/Preferences.md`, `docs/dev/Diagnostics.md` |
| L47 | Specs | new | AGENT | `Core_RedSalamanderMonitor.md` keeps about 40 KB of the removed viewer's requirements as live-sounding prose with nonexistent paths. | `Specs/Core/Core_RedSalamanderMonitor.md` |
| L48 | Packaging | pre-existing | PLAUSIBLE | MSI and MSIX harvest the whole build output folder, so build receipts, toolchain identity and test `*.DxUi.json` sidecars can ship; the ZIP uses an allow-list. | `Installer/msix/RedSalamanderInstaller.wapproj:76`, `Installer/msi/build-msi.ps1:168` |
| L49 | ViewerImgRaw | unknown | NOT DIAGNOSED | ARM64 Release only, flaky (retry passed): an oriented WIC TIFF reports a completed decode but no image (`source=0x0`). Candidates: decode-complete published before the image is installed, or ARM64 `/GL` codegen. Needs an isolated repeat. | `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:5054`, `Plugins/ViewerImgRaw/ViewerImgRaw.cpp:884` |

## Checked and found sound

So nobody re-flags them: the plugin ABI (`FileSystem.h` vtables unchanged, only a new interface, tail-extended options with exact-size check, V1/V2/V3 mutation-result prefixes); `SettingsStore` migrations and schema; `DirectoryRecordCursor`; `SynchronousIoCancelWatch`; `PerfJsonl`; `NaturalSort` (strict weak order, tested empirically); `WindowMessages.h` id uniqueness; every satellite `.rc` placeholder token; `TakeMessagePayload` call sites and payload drains in all viewers; File Operations reducer, conflict broker and arbiter protocol; DecisionMirror (no cross-task decision leak); Issues pane/projection consistency; leftover cleanup cannot delete by itself (ordinary confirmed Delete, Cancel default); Local delete/overwrite/Move/Recycle; S3 test-latency code fully inside `ENABLE_TESTS`; Microsoft/Google Drive diffs are formatting only; ViewerWeb's three merge-day lifetime commits are refcount-balanced and test-only; ViewerVLC worker-owned media switch; ViewerMenuHost ownership; sandbox root authorization, no-follow deletion, owned-VHDX formatting guards, no kill-by-name; clang-format wheel hash pinning; skill trims keep every AGENTS.md rule and all links resolve; FileOps self-test family registration (every scheduled step is in a family); Fold-Changelog `-Check`, Test-Changelog and Get-SpecInventory are green on `main`; the solution, all 114 project include lists, installers and release workflows carry no dangling RedSalamanderMonitor reference.

## Red scheduled workflows: diagnosis

| Workflow (run) | Root cause | Class | Owner finding |
|---|---|---|---|
| Ghostty runtime status (40665991855) | Checkout of 98+ over-long paths without `core.longpaths`; the upload failure is secondary. | repository/CI | H4 |
| Nightly Flake Convergence (40665816743, 40477285880 on `main`; 40691121033 on the repair branch) | On `main`: Checkout (H4), then the upload scans `D:\` (M32). On the repair branch, after those fixes: `BLOCKED ATTESTATION` from the `-SkipBuild` receipt check (M31). | CI/harness | H4, M31, M32 |
| Nightly Native Suite (40660649657) | Build passes in all four profiles; the suite fails. Dominant cause is the hosted desktop: 1024x768 display (work-area fixtures cannot fit) and Windows Terminal owning the foreground (about 45 ARM64 menu/popup/foreground cases). Harness: H2 (ARM64 restoration identity), M28 (x64 foreign-foreground restoration), `theme_cycle_overlay_*` (animations off on hosted OS), `R4A19_DiscoveryIndependentVolumes` (timing gate / no alternate volume). Product, also failing at the baseline: M29, M30, M33. Flaky ARM64 Release: L49. | runner + harness + 3 product | H2, M28, M29, M30, M33, L49 |
| AddressSanitizer weekly (40484804050) | ARM64: the same 59-case set as the nightly plus `Phase9_ConflictPrompt_ApplyToAllUiCache` (0 us convergence wait, unclassified). x64: the FileOps suite printed nothing for 4 h 10 min until the 6 h job limit; the native suite step has no step-level timeout. Hang versus budget exhaustion is not diagnosed (artifacts not downloaded). | runner + harness + undiagnosed | X1 |

The in-flight `ej/ci-runtime-repair-20261028` branch already patches the checkout and upload steps
of `ghostty-runtime-status.yml` and `nightly-flake.yml`; it does not cover `release.yml` and
`winget-release.yml` (H4), MSIX (H3), the harness predicates (H2, M28) or the receipt check (M31).
Whether interactive cases should run on hosted runners at all is Q11.

## Simplification and architecture opportunities

Ordered by value. Each item is behavior-neutral unless it names the finding it also fixes.

| ID | What goes | Size | Risk | Also fixes |
|---|---|---|---|---|
| S1 | Replace the ten per-message host marshals (`Pending*` structs, `SendMessageW` blocks, the counter-token registry, three `kHost*` ids) with one UI-thread call marshal over `SendMessagePayload`/`TakeMessagePayload`, the pattern `AlertOverlayWindow` already uses for UIA (`AlertOverlayWindow.cpp:195-223`). A new service can then not be added unrouted. | about -300 lines in `HostServices.cpp`, `WindowMessages.h`, `FolderWindow.cpp` | medium | H1, M24 |
| S2 | One validated `IFilesInformation` walk: use `Common::Plugins::PackedFileInfoCursor` in `FolderWindow.LeftoverCleanup.cpp`, `FolderView.Enumeration.cpp`, `NavigationView.Menus.cpp` (and check `State.cpp:2896-3040`), plus a shared child-name validator. | about -120 lines | low-medium (FolderView loop is hot; keep it non-allocating) | M14 |
| S3 | Preferences: replace the hand-copied predicate with `Common::Settings::HasNonDefaultFileOperationsSettings`. | -12 lines | low | M3 |
| S4 | Local provider: build stage/backup names with `Common::Paths::BuildUniqueSiblingName`; delete the dead pre-staging overwrite-removal cluster (`CollectDirectoryChildrenVerified`, `RemoveDirectoryRecursiveNoFollow`, `RemovePathForOverwrite`, `FileSystem.FileOps.cpp:4977-5085`); one owned-stage create/abort helper instead of three; one bound-source pump instead of two; fold `StripWin32ExtendedPrefix`/`NtPathToWin32Path`/`Win32PathToNtPath`/`IsExtendedUncPath` into `Common::Paths::ClassifyWindowsPath`. | about -350 lines | medium (data path; run the FileOps families) | M5, M6, M7, L36 |
| S5 | Popup: delete the dead footer/task-button chrome (`DrawFooterAutoDismissControl`, `DrawCheckboxBox`, four unreachable `OnActivatedHit` branches, `_footerNewTasksText`, duplicate AUTODISMISS strings); move `HitTest` under `ENABLE_TESTS`; one status-icon and one ETA helper; carry rate fields in `TaskSnapshot` so a tick builds one snapshot, not two. | about -350 lines | low (S4 of the popup report: medium) | L28 |
| S6 | ViewerSpace: delete never-written layout state (`_layoutCandidateCache`, `_layoutWorkspaceByDepth`, `_autoExpandedOtherByNode`, `_lastRectsByNode` and their types), the dead `_modeCombo` branch and captions, the dead "State" column, 30 unreferenced `IDS_` strings x 5 languages; one `GuardedInvoke` + base class for the two child-host skeletons; move the 575 lines of `#ifdef ENABLE_TESTS` `WndProc` cases to `ViewerSpace.Debug.cpp`; one `ScanProgressState` struct for capture/restore/reset; one paint epilogue; one join-and-release-gate helper. | about -1,100 lines | low | M19, L6 |
| S7 | FileOps engine/runtime: remove the write-only consent-receipt table; one `IsCancellationStatus`; one consent-fact copy helper; one selection-count/phrase helper for the three confirmation builders; one `TakePopupIfIdleLocked()`; drop dead `_pauseCv` notifies in discovery. | about -200 lines | low | L15, L21 |
| S8 | Tooling: delete `Remove-RSTestSandboxStaleRunDirectories` (no caller, weakest guard) and the unused second receipt mechanism in `ScopedTesting.psm1` (about 135 lines); verify the RedMonitor client with MSBuild `GetFileHash` instead of a `pwsh` per project; restrict `format-all.ps1` to the CI path set and the pinned formatter (removes a 90-line resolver). | about -300 lines | low | M25, L42 |
| S9 | Plugins: guard the MTP fake backend with `ENABLE_TESTS`; one `ViewerMenuHost` attach-and-log helper for Web/Text/PE/ImgRaw (only ImgRaw logs a non-hosted outcome today); Dummy uses `PackedFileInfoBuffer::Build`; drop unused `IDS_VIEWERVLC_DETAILS_PATH_NOT_LOCAL_FILE`. | about -1,450 lines from the Release MTP DLL, -150 elsewhere | low | L33 |
| S10 | Split the four god files along their existing seams: `FolderWindow.FileOperations.State.cpp` (25,543 lines; e.g. the 6,000-line cross-file-system bridge, per-item scheduler, batch rename), `FolderWindow.FileOperations.Popup.cpp` (17,666; `Render` alone is 3,400 lines: split the card painter by card kind), `Plugins/FileSystem/FileSystem.FileOps.cpp` (14,507), `Plugins/ViewerSpace/ViewerSpace.cpp` (14,092). | mechanical | medium (large diffs; do it with families green, one file per change) | review-ability, compile time |
| S12 | CI and RedMonitor packaging: one composite "Windows checkout" action instead of eight copies of the long-path step; ship one license file instead of the byte-identical `ThirdParty/RedMonitor/LICENSE.txt` copy; keep one of the two `-text` rules for `ThirdParty/RedMonitor`; retire the `Restore-RedMonitor.ps1` / `RedMonitorDependency.psm1` shims if the inventory allows; delete the dead Monitor constants and stale list entries (L46). | about -100 lines | low | H4, L46 |
| S11 | Search: `MatchNamePattern`/`MatchFileContent` exist twice (`Plugins/FileSystem/FileSystem.Search.cpp:1301`, `RedSalamander/SearchFallbackEngine.cpp:509`, pre-existing); a header-only shared matcher in `Common/` would keep plugin and fallback results identical. | about -80 lines | low | none |

## Remediation slices

Execute in order; each slice is one reviewable change with its own focused validation. Respect the
scoped-testing policy (`Test-Changes.ps1 -Explain`, affected execution, `-Mode PrePush`) and record
each change as a `Changes/` fragment.

1. **Release and qualification blockers (H2, H3, H4, M28, M31, M32, M34, X1).** Coordinate with the
   owner of `ej/ci-runtime-repair-20261028` first so the work is not done twice. Add
   `'$(MSBuildProjectExtension)'=='.vcxproj'` to `ValidateMsvcToolsVersion` and a Pester check that
   evaluates the `.wapproj` (H3). Put the long-path checkout in one composite action used by every
   Windows job, with a policy test, or shorten the evidence paths (Q12) (H4). Fix the restoration
   predicate (drop the same-thread requirement, evaluate the no-write shortcut first, print GA_ROOT
   and focus facts) and add a bounded settle plus diagnostics to foreign-foreground restoration
   (H2, M28). Make the receipt check name its failed predicate and restore or exempt the DxUi source
   for the portable consumer (M31). Guard the Flake upload (M32). Pass the overlays in the inline
   vcpkg step and add `Build/vcpkg-ports` to the cache key (M34). Add a step timeout to the native
   suite and diagnose the silent x64 ASan FileOps run. Exit: scheduled workflows green on one `main`
   commit; Fresh Full passes for x64 Debug, Release and ASan Debug on that identity; `build.ps1 -Msix`
   produces a package.
2. **Host services (H1, M24, S1).** Add worker-thread self-tests for every `IHost*` entry point first
   (they must fail today for the three unrouted services), then collapse the marshaling. Exit: a
   worker-thread `SetConnectionSecret` round-trips through `GetConnectionSecret`; forged `lParam`
   values are ignored; `docs/dev/PluginHostModel.md` corrected.
3. **User-visible correctness quick wins.** M1 (pass `FILESYSTEM_FLAG_RECURSIVE` for viewer
   Delete/Recycle, with a host-side non-empty-folder test), M3/S3 (predicate), M13 (Running name),
   M10 (admitted item total in the interrupted card), M19 (delete the `_modeCombo` branch), M20
   (single count of workspace bytes), M21 (share root has no parent), L5 (translate or reuse the
   existing string).
4. **File Operations concurrency (M2).** Apply the `WakePauseWaiters` idiom (take the waiter's
   mutex after the flag store, before `notify_all`) at every listed site; add a stress self-test that
   hammers Pause/Resume/Cancel against a scheduler wait. Run the full FileOps families.
5. **Local provider (M4, M5, M6, M7, L36, L37, S4).** Data path: run every FileOps family and the
   Local plugin contract tests. Add cases: Overwrite of a Hidden file keeps Hidden; 230-character
   leaf Overwrite succeeds; `\\?\`-target symlink Preserve round-trips; copying a file being appended
   succeeds; the stage name contains no PID/TID.
6. **Leftover cleanup and plugin buffers (M14, M15, M16, M17, S2).** Shared validated walker, child
   name validation, `shared_ptr` scan context with a bounded destroy wait, an "unreadable folders"
   status line, conservative pre-check (Q5).
7. **File Operations UX (M9, M11, M12, L17, L18, L21).** Breadcrumb writer identity and once-per-
   process load (Q3), inline-rename pre-bind terminal results (Q8), Escape-binding visibility rule
   (Q4), notices exempt from the summary cap, plural strings.
8. **MTP (M8, L32, L33).** Memo only after a successful direct fallback, no ACCESS_DENIED, reset on
   reconnect, `SupportsAtomicWriterCommit` honest while set (Q9); guard the fake backend.
9. **FolderView and ViewerSpace resilience (M18, M22, L1-L14).** Per-generation stop source for
   enumeration (Q6); ViewerSpace reaper for cooperative workers instead of immediate detach.
   Search and throughput failures that the nightly reproduces (M29, M30, M33) belong here too:
   tracked-root traversal for the journal-unavailable rebuild, the pipe-busy connect policy, and
   per-item pool sizing.
10. **Tooling (M25, M26, M27, L40-L42, S8).**
11. **Remaining Low items and simplifications S5-S11**, one area per change.
12. **Repo hygiene (L43, L44)** after the owner answers Q12.

## Questions for the owner

Answers change what the slices do; until answered, the slice keeps the current behavior.

1. **Release gate.** Is "production-ready" defined as green scheduled workflows plus a Fresh Full
   in x64 Debug, Release and ASan Debug at one `main` identity (slice 1 exit)? Do native ARM64
   execution and interactive suites belong to that gate or stay separate handoffs (as I48 records)?
2. **Host services (H1).** Should a host service ever report success for an unrouted call, or should
   the marshal fail closed (`E_NOTIMPL`) so a future gap is visible?
3. **Move breadcrumbs (M9, L17).** Bind each record to its writer (process id + start time, or a
   per-process session id) and load notices once per process? Keep notices until dismissed,
   exempt from the 24-summary cap? Should `Finalize` try the removal before the Terminal write so a
   failed write cannot produce a false "interrupted" warning?
4. **Popup Escape (M12).** Should the sole-decision Escape fallback require the decision card to be
   visible (or scroll it into view when the decision is published)? Should a Partial result show
   a full bar and "100%" (popup report)? Is a modeless-in-practice Custom speed-limit prompt intended?
5. **Leftover cleanup (M15-M17).** Pre-check by name shape alone, or leave unchecked anything modified
   in the last N minutes (or require a cross-window lease)? When the walk is incomplete, keep Delete
   Selected enabled (report only) or disable it? Is an unbounded UI wait at window destroy for a
   wedged remote walk acceptable?
6. **FolderView Cancel (M18).** Must Cancel abort the provider enumeration (spec `UI_FolderView.md:514`),
   or is "abandon the result, keep the worker busy" intended (then the spec changes)?
7. **ViewerSpace.** Escape in the analysis panel: hide the panel or cancel/close the viewer? File-count
   weighting: should eligibility and squarify order follow the weight? Up at `\\server\share`: disabled
   (assumed)? Should one malformed provider directory (control character, case-only duplicate in a
   case-sensitive folder) fail the whole scan or only that folder? Is permanent unload quarantine after
   a mid-scan close acceptable, or adopt a reaper? Export of an unrepresentable name: fail the export or
   substitute U+FFFD with a flag?
8. **File Operations results.** Inline Rename pre-bind refusals: a terminal "not renamed" result like
   Batch Rename (assumed)? Should skipped and partially failed roots count in the live "N of M items"
   meter? Is the consent-receipt table meant to become an audit record, or delete it?
9. **MTP (M8, L32).** Should `ERROR_ACCESS_DENIED` on the hidden temp count as "device cannot rename"?
   Session-wide memo, or per folder / re-learned after reconnect? Must creates get the same journal and
   replay guarantee as overwrites?
10. **Local provider.** Does `GetFileInformationByHandleEx(FileIdInfo)` succeed on FAT32/exFAT/UDF
    media? If not, native Copy from and Overwrite onto USB sticks fail like the NAS case; this needs a
    manual check with a real stick. Should one malformed directory record fail the whole listing
    (`DirectoryRecordCursor`) or be skipped with a Warning? Should the spec's "no predictable stage
    name" rule govern the Local provider (assumed yes, M5)?
11. **Tooling and hosted runners.** Retention and reclamation for
    `X:\RedSalamander.Perf\evidence\completed-runs` (kept forever today, including detached VHDX
    files)? Should interactive Commands/Viewer cases run on hosted runners (1024x768, Windows Terminal
    in the foreground on ARM64 Windows 11), or move to a self-hosted desktop? On a hosted runner, is an
    unrestorable foreign foreground with unchanged input a failure that aborts the Viewer suite, or
    inconclusive evidence? Should `-NonInteractive` on legacy suites exit 2 as documented? Should MSI and
    MSIX carry build receipts and test sidecars, or use the ZIP allow-list (L48)?
12. **Public repository content (H4, L43).** Must the repository stay cloneable on stock Windows git
    (`core.longpaths=false`)? If yes, the 607 long evidence paths must be shortened or moved. Should the
    public repo carry the 19,500 evidence files under `Specs/TestRuns`, `advisor-plans/` and
    `.codex/config.toml` (which grants agents network access)? Should the private NAS hostname be
    replaced by a fixture name in published plans and evidence?
13. **Shell.** Should direct-mutation commands (Change Attributes, Create Directory, Edit/New File,
    Pack/Unpack, Make File List) also be refused while an accepted exit is pending? Can a Compare pane
    host an embedded Terminal (its exit notification goes to the Compare window, which has no handler)?
    Should the Preferences "Monitor" page keep the "RedSalamanderMonitor" name? Should the RedMonitor
    settings path follow the viewer's version instead of the fixed `-7.0` file?
14. **Diagnostics (L38).** Truncate ETW messages over 16,000 code units with a marker, or drop and count?

## Out of scope and stop conditions

- This plan authorizes no push, PR, release, workflow re-run or change in another repository
  (DxUi, RedMonitor). Each slice is executed only after the owner picks it.
- Stop if the drift check shows an in-scope file changed and the cited excerpt no longer matches;
  re-verify the finding first.
- Stop a data-path slice (4, 5, 6, 8) if any FileOps family regresses; preserve the failed run.
- Do not relax timeouts, delete failing tests or weaken focus/sandbox guards to make slice 1 green.

## Done criteria

- Every High and Medium finding is fixed with a failing-first regression case, or explicitly accepted
  by the owner with the answer recorded here.
- Slice 1 exit holds at one `main` identity, with the receipts archived under `Specs/TestRuns`.
- Durable rules found here are merged into their owning specs (`Plugins_PluginAPI.md` marshal rule,
  `FileSystem_FileOperations.md` breadcrumb ownership and stage naming, `UI_FileOperationsPopup.md`
  Escape rule, `UI_FolderView.md` Cancel, `Plugins_ViewerSpace.md` answers, `Core_StartupBootstrap.md`).
- Then move this plan to `Specs/Plans/Done/` through the reviewed Done-admission steps and remove its
  index row.

## Evidence

Per-area reviewer reports and the lead verification ledger were working files for this review and
are not retained in the repository; every finding above carries its own file and line references.
