# Operation Evergreen — Three-Day Diff Review Remediation (2026-07-05)

Remediation ledger for the multi-agent review of everything changed **2026-07-02 → 2026-07-05**:
base `275c04034` → master `da6b438a0` **plus the uncommitted working tree** (Operation Granite
search/MTP slice, 17 files). Code delta reviewed: ~92 files, +8,813/−1,457 lines
(`Specs/TestRuns` archives excluded). Full evidence for every item lives in
**`Specs/Reviews/ThreeDayDiff-2026-07-05-Findings.md`** (referenced below as *F#n*); this plan is
the execution ledger — work items, fixes, gates, sequencing.

Method: 13 finder agents (9 area reviewers + architecture/concurrency/merge-integrity lenses) →
81 raw findings → 69 after dedup → adversarial verification (bugs verified by two independent
verifiers: correctness + reachability). 61 fleet-confirmed + 3 manually re-verified = **64
confirmed findings**; 3 refuted (recorded in the findings doc appendix — do not re-report them).

> **Planned-at anchors.** Master `da6b438a0`, review base `275c04034`, working tree of
> 2026-07-05 (dirty: Granite search/MTP slice uncommitted). All `file:line` anchors are
> verifier-corrected against that tree. Re-check drift with `git diff <anchor-file>` before
> executing any item — this ledger does not auto-track the tree.

> **Coordination caution.** Tracks A (MTP), C5–C7 (search), and E9–E10 (Mtp/SearchAndIndex
> selftests) touch files that were changed by Operation Granite. Granite is now committed and
> archived under `Specs/Plans/Done/` (main ledger plus
> `Operation_Granite_ContinuationBaton_2026-07-05.md`), so re-check drift against current
> `master` before executing these items. Track B touches DxUi files owned by the UIA baton
> session — see the routing table.

## Routing table (single-owner rule)

| Evergreen item | Routed to | Why |
|---|---|---|
| EV-B2 (ungated a11y rebuild triggers, F#29) | `DxUi_Uia_ContinuationBaton_2026-06-29.md` (Granite GR-A1 coalescing pass) | That baton already owns the dirty-flag/debounce + `UiaClientsAreListening` gate. F#29's inventory of the *new* trigger sites (mouse-up, Grid/Tree mutators, Control setters) is required input for its coalescing design — hand it over, do not implement twice. |
| EV-B3 (13× copy-pasted refresh block + public `Grid::RefreshAccessibilitySnapshot`, F#42) | same baton | The consolidation point (one protected `Control::RefreshAccessibilitySnapshot()`) is exactly where the baton's dirty-flag will live; consolidating first makes the gate a one-line change. |
| Systemic flake-harness work surfaced by Track E | `Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md` | Evergreen owns only the *defects in code added this window*; retry/quarantine/per-case-reset stays with the stabilization plan. |

Everything else is owned here. No other plan tracks these findings (verified against the
2026-07-02 WIP reorg routing map; the MTP items are defects **in** the landed Granite GR-P1/GR-P2
work, not duplicates of it).

---

## Track A — MTP WPD cache & lifecycle (P0: plugin viability)

The new session/path cache in `FileSystemMtp.Device.cpp` traded the old open-session-per-operation
resilience for speed without adding any recovery path. A1–A4 are the viability core; do them as
one coherent pass over the cache layer, in this order (A3's COM-lifetime decision changes what A1
eviction must do).

### EV-A1 [HIGH] Cache invalidation on operation-phase failure (F#4, F#6, F#7)
`FileSystemMtp.Device.cpp:1427` (EnumerateDirectory), `:1500` (ReadFile), `:1641` (DeleteItem),
`:1678` (RenameItem), `CreateFileReader`/`UploadFileObjectCached` `RETURN_IF_FAILED` sites,
`GetOrOpenSession:1997-2003`.
- **Problem.** `FailAndMaybeInvalidateCaches()` is wired only into the RESOLVE phase; every
  operation-phase error returns raw. `GetOrOpenSession` hands back cached
  `IPortableDevice`/`IPortableDeviceContent` with no liveness check. After a device replug (or WPD
  hiccup), a cached path resolves with **zero device I/O**, the operation fails promptly on the
  dead session, nothing is evicted, and every retry is byte-identical — the natural user action
  (refresh a visited folder) loops forever. Core-side recovery (`AbandonBackendSessionLocked`)
  triggers only on watchdog TIMEOUT or manual Disconnect; a prompt failure never trips it.
  Additionally `ShouldInvalidateAllCachesOnFailure` whitelists `ERROR_FILE_NOT_FOUND`/
  `ERROR_PATH_NOT_FOUND`, and `EnumerateDirectory` fetches fresh items without refreshing
  `_pathCache` — listing and cache are two divergent sources of truth.
- **Fix.** (1) Route every operation-phase failure through `FailAndMaybeInvalidateCaches` — at
  minimum evict the `_sessionsByPnpId` entry and the pnpId's path-cache subtree on any
  non-whitelisted HRESULT. (2) On missing-object errors from a **cache-resolved** objectId, evict
  that specific entry + descendants and retry once with a fresh resolve. (3) Write/refresh
  `_pathCache` child entries from every successful `EnumerateDirectory` result.
- **Verify.** RED-first selftest on the WPD-cache fixture (after EV-A9/EV-TA1): resolve a path,
  simulate session death (fixture hook), assert next operation reopens a fresh session instead of
  returning the same error twice. Manual: browse phone → replug → refresh folder recovers.

### EV-A2 [HIGH] Cached item metadata freshness for size-sensitive ops (F#5)
`FileSystemMtp.Device.cpp:1497` (ReadFile → `ReadPortableDeviceStream` expectedSize), `:1145`
(`WpdStreamBackendFileReader::GetSize`), `:1582` (GetFileSize), `TryResolveCachedPath:2042`.
- **Problem.** Cache hits return the `MtpItem` captured at first resolution; no TTL, no refresh on
  success. `ReadFile` passes the stale `sizeBytes` as `expectedSizeBytes` and
  `ReadPortableDeviceStream` returns `ERROR_CRC` on any mismatch (lines 1018-1022) — which per
  F#4 doesn't invalidate — so a file modified on the device (in-progress recording, download)
  fails **permanently** even though the live listing shows the new size. Stale `GetSize` also
  feeds the host move flow's size comparison (`FolderWindow.FileOperations.State.cpp:8329`) →
  spurious `ERROR_PARTIAL_COPY`.
- **Fix.** Re-fetch object properties for files before read/transfer (or drop the cached entry
  when it's about to be used size-sensitively); invalidate the entry whenever the size check
  fails. EV-A1's enumeration-refresh also narrows the window.
- **Verify.** Fixture test: cache a file, change its size via fixture, read → must succeed with
  new size (RED before fix: ERROR_CRC).

### EV-A3 [HIGH] COM lifetime: init once per worker-thread lifetime (F#8, F#13)
`FileSystemMtp.Device.cpp:44-65` (per-call `ComInitialization`), `:2401` (`_sessionsByPnpId`),
`:1124-1233` (`WpdStreamBackendFileReader`), `FileSystemMtp.Core.cpp:2111-2180`
(`MtpBackendReader::GetSize/Seek/Read` — no COM guard at all).
- **Problem.** Cached WPD COM interfaces (sessions, streams) outlive the per-command
  `CoInitializeEx`/`CoUninitialize` scopes on the queue worker. If a `CoUninitialize` releases the
  process's last MTA reference, the apartment tears down under the cached pointers →
  `RPC_E_DISCONNECTED`/AV on the next command. Reader IStream calls are the only WPD path with
  **no** `ComInitialization` whatsoever.
- **Fix.** `CoInitializeEx(MTA)` once in `MtpBackendCommandQueue::WorkerMain`, `CoUninitialize` at
  worker exit; drop the per-call scopes (or keep as no-op fallback). Alternative:
  `CoIncrementMTAUsage` cookie held by `WpdMtpBackend`. All backend and reader calls already
  funnel to the worker, so one init covers everything.
- **Verify.** Existing MTP selftests still green; add an assert (debug) that WPD calls observe
  COM initialized.

### EV-A4 [HIGH] Recreate the backend worker on re-Initialize (F#3)
`FileSystemMtp.Core.cpp:2241` (`AbandonBackendSessionLocked` guarded recreate is dead code —
both callers set `_disconnected = true` first: `:2303-2309`, `:3665-3672`), `Initialize:2582`.
- **Problem.** After any watchdog trip or Disconnect, `_backendWorker` stays null forever. The
  documented re-arm path (`Initialize` resets `_disconnected`; `FolderWindow.FileSystem.cpp:5135`
  re-Initializes the cached pane instance) leaves the instance permanently returning `E_FAIL`.
- **Fix.** In `Initialize`, recreate `_backendWorker` when null and `_backend` is set (or drop the
  `_disconnected` guard in `AbandonBackendSessionLocked`). Decide explicitly what outstanding
  `MtpBackendReader`s bound to the abandoned backend do — that decision is EV-A5.
- **Verify.** RED-first: fixture watchdog trip → re-Initialize → next command must succeed.
  Coordinate with the existing `mtp_backend_command_worker_is_reused` contract case.

### EV-A5 [LOW→raised by A4] Tie MtpBackendReader to its backend generation (F#31)
`FileSystemMtp.Core.cpp:2038/2111/2139/2176`.
- **Problem.** Readers capture raw `FileSystemMtp*` + shared `IMtpBackendFileReader` and submit to
  the owner's **current** `_backendWorker`. Today masked by the A4 bug (calls fail fast on the
  null worker); the moment A4 lands, a stale pre-trip reader would run its old backend's IStream
  on the NEW queue thread concurrently with the quarantined old worker — same WPD session, two
  threads. **Must land with EV-A4.**
- **Fix.** Store the `shared_ptr<IMtpBackend>` (or a session generation counter) in the reader;
  fail `ERROR_DEVICE_NOT_CONNECTED` when it no longer matches.

### EV-A6 [MEDIUM] Overwrite-journal absent-cache cross-instance race (F#12)
`FileSystemMtp.Core.cpp:1392` (`RecordOverwriteJournalIntent` invalidates at entry `:1134`, writes
`:1160`, no re-invalidate after write — unlike `WriteOverwriteJournalEntry:1086-1090`).
- **Problem.** `g_absentOverwriteJournalIdentities` is process-global; journal ops are serialized
  only per instance. Two panes on one device: B invalidates → A's `ClearOverwriteJournalIntent`
  deletes the file → B writes its journal → A's `MarkOverwriteJournalAbsent` caches "absent" while
  B's live journal exists → crash-recovery replay skipped.
- **Fix.** Re-invalidate after the successful write in `RecordOverwriteJournalIntent`; make
  mark-absent conditional on no intervening invalidation (per-identity generation counter under
  `g_overwriteJournalAbsentMutex`). Also normalize the cache key with the same `towlower` folding
  as `StableDeviceHash`.

### EV-A7 [LOW] RenameItem overwrite: guard directory sources (F#18)
`FileSystemMtp.Core.cpp:3352`.
- **Problem.** The new allow-overwrite RenameItem never checks SOURCE attributes. Directory
  source onto existing file: real WPD fails midway with `ERROR_NOT_SUPPORTED`; the fake backend
  happily replaces the file with the directory — backend divergence means tests pass where
  hardware fails.
- **Fix.** `GetAttributes(commandSource)` in the overwrite branch; fail fast
  (`ERROR_ACCESS_DENIED`/`ERROR_NOT_SUPPORTED`) for directory sources; apply the same guard to the
  CopyItem/MoveItem overwrite lambdas. Folds naturally into EV-A12's shared helper.

### EV-A8 [LOW] noexcept constructor starts a jthread (F#19)
`FileSystemMtp.Core.cpp:2212`, `FileSystemMtp.h:90-91`.
- **Problem.** `std::make_shared<MtpBackendCommandQueue>` in a `noexcept` ctor: thread exhaustion
  or OOM → `std::terminate`, while the surrounding factory deliberately uses
  `new(std::nothrow)`/`E_OUTOFMEMORY`.
- **Fix.** try/catch around queue creation (leave `_backendWorker` null; `RunBackendCommand`
  surfaces failure) or create the worker lazily on first Submit. Note interaction with EV-A4's
  "recreate when null" logic — same code path.

### EV-A9 [LOW] Selftest WPD-cache backend: null-deref guards on the four unguarded mutators (F#20)
`FileSystemMtp.Device.cpp:1890` (selftest `OpenDeviceSessionForBackend` returns S_OK with null
sessions), DeleteItem `:1636`, RenameItem `:1672`, Copy/Move `:1711/1761/2372`.
- **Fix.** Add the existing `if (_selfTestWpdCacheBackend) return E_NOTIMPL;` guard (present in
  WriteFile/CreateDirectory at `:2283/:2336`) to DeleteItem/RenameItem/CopyItem/MoveItem, or
  null-check `content` in the ById helpers. Superseded if EV-A10 lands first.

### EV-A10 [LOW/arch] Extract the selftest fixture from the production backend (F#32, F#33)
`FileSystemMtp.Device.cpp:1833` (~120 lines of fixture data + `if (_selfTestWpdCacheBackend)`
forks in production hot paths; only the factory is `#ifdef _DEBUG` at `:2418`).
- **Problem.** A second fake device is compiled into release `WpdMtpBackend`; partial guard
  coverage caused EV-A9; `GetInfo()` reports `liveWpd=true` for the fixture, confusing Core's
  abandon logic. The project already has `FakeMtpBackend` behind the `IMtpBackend` seam for
  exactly this role.
- **Fix.** Extract the caching/resolution layer over a small device-ops interface; implement the
  fixture as a separate implementation (like FakeBackend); report `liveWpd=false` for it. Interim
  cheap step: `#ifdef _DEBUG` the fixture members and branches.

### EV-A11 [LOW] `_backendThreadIdsOverflow` data race (F#21)
`FileSystemMtp.FakeBackend.cpp:868` (read under `_mutex`; writes at `:905` under
`_threadStatsMutex`).
- **Fix.** Locked accessor (`BackendThreadIdsOverflow()` taking `_threadStatsMutex`, or one
  accessor returning both values) used by `GetItemProperties`.

### EV-A12 [MEDIUM/simplification] MTP dedup & dead machinery (F#43, F#48, F#49, F#50)
- Extract one helper for the ~40-line overwrite orchestration duplicated across
  CopyItem/MoveItem/RenameItem (`Core.cpp:3103/3199/3320`) — carries EV-A7's guard once.
- Delete `deviceIoMutex` (only ever locked by the queue's single worker, `Core.cpp:2001`; the
  queue *is* the serialization point — keep a comment saying so).
- Inline `WriteOverwriteJournalReplayAttempt` (one-line wrapper, single caller `Core.cpp:1515`).
- Merge `WpdMemoryBackendFileReader` (`Device.cpp:1027-1122`) and `FakeMtpStreamingReader`
  (`FakeBackend.cpp:253-360`) — byte-identical seek/overflow logic — into one
  `MemoryBackendFileReader` in `FileSystemMtp.Internal.h` with an optional per-read hook.

### EV-TA1 [MEDIUM/test] `mtp_wpd_session_and_path_cache_reuse` release-build skip guard (F#15)
`CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp:2220`.
- **Fix.** Add the sibling cases' guard:
  `if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND)) return state.Skip(...)` — both
  selftest factory exports are `#ifdef _DEBUG` (`Factory.cpp:119-181`), so the case currently
  hard-fails the whole suite against a release plugin.

---

## Track B — DxUi accessibility & text input

### EV-B1 [HIGH] O(selection²) offscreen-selected-row snapshot loop (F#1)
`DxUi.Accessibility.cpp:1152`; model scan `FindFilesWindow.cpp:1084-1091`; triggers
`DxUi.Grid.cpp:3447/3649/3269`, `DxUi.WindowHost.cpp:2567`.
- **Problem.** Per selected row: linear `FindRowByStableId` + `std::find_if` over `gridRows`
  (grows per offscreen row → quadratic) + `GetCellData`/`BuildGridCellAccessibleText` per visible
  column — synchronously on the UI thread, on every selection change; Ctrl+A in a large Find Files
  result set freezes the UI.
- **Fix.** Cap materialized offscreen selected-row records (e.g. first 100), build a hash set of
  visible rowIds once, and/or resolve offscreen row names lazily in the GridRow provider. Land
  independently of the routed coalescing work (which reduces *how often*, this fixes *how much*).
- **Verify.** Perf selftest: Ctrl+A on a 10k-row grid, budget the snapshot rebuild.

### EV-B2 → ROUTED to `DxUi_Uia_ContinuationBaton_2026-06-29.md` (F#29 — ungated rebuild triggers)
### EV-B3 → ROUTED to `DxUi_Uia_ContinuationBaton_2026-06-29.md` (F#42 — 13× refresh copy-paste + public Grid refresh)
Hand both findings' site inventories (in the findings doc) to the baton owner; its GR-A1
dirty-flag/coalescing pass must cover the **new** trigger sites or the gate will be incomplete.

### EV-B4 [HIGH] TSF deactivation use-after-free window (F#2)
`DxUi.NativeTextInput.cpp:842` (Disconnect/Detach moved AFTER `Pop(TF_POPF_ALL)`; store holds raw
`_host`/`_control` during Pop; `TextField::~TextField` never notifies the host —
`DxUi.TextInput.cpp:925`; staleness found only lazily via `PruneStaleInteractionState`
`WindowHost.cpp:3842-3845` / `OnKillFocus:3754`).
- **Fix.** Before Pop, verify the control is still alive/in-tree (`ControlBelongsToTree` or a
  lifetime token); if stale, use the old Disconnect-first ordering; keep the new keep-attached
  ordering only when host+control are known-live. Robust version: store holds a
  `std::weak_ptr` lifetime token checked in every entry point that derefs `_control`.
- **Verify.** Extend `DxUiTests.NativeTextInput` destroy-during-deactivate coverage (the existing
  `...DeactivateNativeTextInputBeforeDestroyWindow` test is the anchor).

### EV-B5 [LOW] Dead machinery in DeactivateNativeTextInputTsf (F#45)
`shouldReleasePreviousFocusDocumentAfterPop` is a no-op (member reset unconditionally at `:876`);
Disconnect (`:887`) + Detach (`:895`) are the same call (`DxUi.TextStoreACP.cpp:177-180`). Delete
the bool + conditional; keep a single Disconnect. Fold into EV-B4's edit.

### EV-B6 [LOW] TSF trace scaffolding cleanup (F#46)
18+ `TraceNativeTextInputTsfStep` sites re-read the env var and reopen the file per line, burying
the ordering-sensitive teardown logic. Remove, or collapse to one scoped tracer (cached env
lookup). Do after EV-B4/B5 so the diff reviews cleanly.

### EV-B7 [MEDIUM] Secure wipe on the layout-cache replacement path (F#11)
`DxUi.SingleLineTextEditing.cpp:648`.
- **Problem.** `secureCacheText` wipes only the three failure/clear paths; the common path
  (`cache->text = std::wstring(text);`) frees the old buffer un-wiped / leaves a stale tail on
  shrink. Revealed-password text cycles through this on every keystroke
  (`GetDisplayText` returns `_text` for `PasswordRevealMode::Visible`, `DxUi.TextInput.cpp:3107`).
- **Fix.** When `secureCacheText`, `SecureWipe::SecureClear(cache->text)` before assigning, and
  assign via buffer-exact copy so no stale tail survives capacity reuse.
- **Verify.** Strengthen the weak contract assertion at the same time — that is EV-F6 (F#39).

---

## Track C — Settings, hot reload, connections, search

### EV-C1 [HIGH] SettingsHotReload::Start blocks UI thread 2s and permanently disables hot reload on first run (F#10)
`SettingsHotReload.cpp:518`; watcher signals ready only after `FindFirstChangeNotificationW`
succeeds (`:88-103`); settings dir is created only on first save
(`Common/SettingsStore.cpp:384`).
- **Fix (pick one, first preferred).** (a) Signal `readyEvent` as soon as the thread starts —
  "alive and retrying" is success; (b) on `WAIT_TIMEOUT` leave the watcher running instead of
  `Stop()`; (c) `CreateDirectoryDeep` the settings dir before starting the watcher.
- **Verify.** First-run scenario (no settings dir): Start returns fast, hot reload works after
  the dir appears.

### EV-C2 [LOW] Capture GetLastError before Stop() in the WAIT_FAILED path (F#26)
`SettingsHotReload.cpp:527`. `const DWORD lastError = GetLastError();` immediately after
`WaitForSingleObject`.

### EV-C3 [MEDIUM] ConnectionManager: validation-fallback rename desyncs `_selectedConnectionName` (F#62)
`ConnectionManagerWindow.cpp:4160-4163` (refresh only when `GetSelectedModelIndex()` has a value),
`:3082` (fallback rename), `:3321`/`:4175` (stale name consumed), `:4117` (Close clears the name
*before* validating, silently dropping the same rename).
- **Fix.** After `TryValidateAndNormalizeConnectionProfiles()`, refresh `_selectedConnectionName`
  via `ResolveEditedModelIndexForValidation()` (or have validation return the resolved index).
  Decide the Close-path behavior explicitly (apply-or-discard, but not silently).
- **Verify.** Selftest: clear grid selection, rename in editor, Connect → saved settings and
  `_modalResult->connectionName` must agree.

### EV-C4 [LOW] Delete the unreachable `case WM_NCDESTROY:` (F#51)
`ConnectionManagerWindow.cpp:4730` — WM_NCDESTROY is intercepted at the top of WindowProc
(`:4692-4696`).

### EV-C5 [MEDIUM] Search: per-root service rejections poison the global 5s cooldown (F#63)
`FileSystem.Search.cpp:2803-2807` (cooldown armed for **every** fallback-candidate HRESULT),
`:2660-2661` → `:2705-2707` (subsequent searches skip service), `:2711-2714` (mislabeled
`SERVICE_UNAVAILABLE` warning on unrelated queries).
- **Problem.** The GR-13 per-request fallback (root refused → degrade to INDEX/SCAN) is
  intentional and test-covered; the collateral is that a root rejection from a **healthy** service
  arms the instance-wide cooldown, degrading unrelated searches for 5s and labeling the service
  unavailable.
- **Fix.** Arm the cooldown only for transport/connect-class failures (pipe missing, timeout) —
  not for per-request validation rejections; or use a distinct warning flag
  (`ROOT_REJECTED_BY_SERVICE`) for the per-root case. **Do not** change
  `IsServiceFallbackCandidate` itself (see refuted finding #2 in the findings-doc appendix — the
  fallback candidacy is deliberate GR-13 behavior with RED/GREEN coverage).
- **Verify.** Selftest: healthy service rejects root A → immediate search on valid root B still
  uses the service backend.

### EV-C6 [LOW] ConnectClientPipe missing-pipe retry loop needs a cancel hook (F#16)
`SearchServiceBroker.cpp:1519`. The env override
(`REDSALAMANDER_SEARCH_SERVICE_CLIENT_MISSING_PIPE_RETRY_MS`, harness sets 5000, cap 30000)
stretched an uncancellable `Sleep(10)` poll from 250ms to up to 30s per connect; `cancelCheck` is
consulted only post-connect. Thread the caller's `CancelCheckFn`/stop event into the loop; return
`ERROR_CANCELLED` promptly.

### EV-C7 [LOW] AuthorizeClientRootRebuildAccess unreachable tail (F#47)
`SearchServiceBroker.cpp:2518`. Loop cannot exit via its condition (root non-empty at entry,
parent-empty check returns inside); the trailing `return HRESULT_FROM_WIN32(lastMissingError);`
is dead and its `ERROR_PATH_NOT_FOUND` initializer never read — misleading in a
security-sensitive walk-up. Use `for(;;)` + delete the tail (or `std::unreachable()`), drop the
initializer.

---

## Track D — FolderView / FolderWindow file operations

### EV-D1 [LOW, wedge] `_pasteShortcutInFlight` stuck forever on lost completion (F#22, F#23)
`FolderView.FileOps.cpp:990` (flag set), `:979` (worker's `PostMessagePayload` result discarded),
`:1048` (only reset site), `:874-885` (subsequent requests queue forever, reporting success).
- **Fix.** On `PostMessagePayload` failure: retry/post a sentinel; plus a UI-side staleness reset
  (timeout or `WM_NCDESTROY` payload draining). At minimum log the failure. One fix closes both
  findings.

### EV-D2 [LOW] Completion notifies DirectoryInfoCache with the pane's *current* `_fileSystem` (F#24)
`FolderView.FileOps.cpp:1032`. Navigating the pane to another plugin mid-work pokes the wrong
provider context (cache is keyed per context, `DirectoryInfoCache.h:139-152`) and the builtin
context that actually changed is never notified. **Fix.** Capture the `IFileSystem` (or provider
identity) in `PasteShortcutRequest`/`Result` and notify that context.

### EV-D3 [MEDIUM] Shortcut-save temp-file name reuse race (F#14)
`FolderViewInternal.h:1215` (`CreateShellPersistTempPath`: `GetTempFileNameW` reserves →
`DeleteFileW` at `:1197` frees the name → `persist->Save` later). Concurrent cross-pane saves or a
second RS instance sharing `%TEMP%`+`rsl` prefix can steal the name → wrong-target shortcut or
spurious failure.
- **Fix.** Don't delete the placeholder (`IPersistFile::Save` with `STGM_CREATE` overwrites it),
  or use a GUID name; ideally create the temp in the destination folder so the final
  `MoveFileExW` is an atomic same-volume rename.

### EV-D4 [LOW/simplification] Issues-pane hide + focus-restore dedup, fix SaveViewState divergence (F#52, F#53)
`FolderWindow.FileOperations.IssuesPane.cpp:889-909` vs
`FolderWindow.FileOperations.State.Diagnostics.Part.cpp:221-237` — same 12-line block, except
OnClose also calls `SaveViewState()` and ToggleIssuesPane does not (toggle-hide silently loses
grid view state). **Fix.** Have ToggleIssuesPane send the pane `WM_CLOSE` (reuse OnClose), or
extract one helper that saves state, hides, restores focus. The divergence is a real (minor)
behavior bug, not just cleanup.

### EV-D5 [LOW/simplification] One prompt-close debug helper instead of three, keep the `message==0` guard (F#54, F#55)
`FolderWindow.FileSystem.Commands.Part.cpp:3863` (has guard),
`FolderWindow.FileSystem.Navigation.Part.cpp:853` (drops guard — posting `WM_NULL` on
registration failure turns a selftest 'failed to confirm' into a misleading full-timeout wait),
`FolderWindow.FileOperations.Popup.cpp:8118`. Hoist a single guarded helper into a shared
ENABLE_TESTS internal header.

### EV-D6 [LOW/simplification] Dropdown focus-restore block ×3 (F#56)
`NavigationView.Menus.cpp:2359` (new copy; existing at `:2116-2124`, `:2627-2634`). Extract
`NavigationView::RestoreFolderViewFocusAfterDropdown()`.

### EV-D7 [LOW] Misplaced `RestoreDeferredFocusAfterLayout()` (F#57)
`Preferences.Dialog.cpp:5221` — sits in the Plugins deferred-action branch; the compare case is
already handled by `LayoutPreferencesPageHost` (`:3788`). Delete the call; also clear
`_deferredFocusAfterLayout` in `CompareDirectoriesPane::DetachDxHosts` so a stale target cannot
survive pane teardown.

### EV-D8 [LOW] FileSystemDummy last-instance cleanup: decrement-then-lock race (F#17)
`FileSystemDummy.cpp:2896`. Take `std::scoped_lock(_mutex)` around the increment in the ctor and
around decrement + `ClearRootsIteratively()` in the dtor so the last-instance check and the roots
clear are atomic vs new-instance construction.

---

## Track E — Self-test correctness (tests that cannot do their job)

### EV-E1 [HIGH] Raw UIA-provider fallback is a silent no-op on the worker-thread paths (F#9)
`Commands.SelfTest.Settings.cpp:5289-5293` (helpers run on a fresh `std::jthread`);
`CreateWindowHostAccessibilityProvider` returns null off the window thread
(`DxUi.Accessibility.cpp:7108-7118`; `ResolveHost` thread gate `:353-357`).
- **Fix (choose).** (a) Marshal provider creation to the window thread via SendMessage test hook;
  (b) relax the ResolveHost gate for the ENABLE_TESTS factory (reads are snapshot-based, actions
  already marshal); (c) drop the raw fallback from worker-side helpers and delete
  `SetWindowHostRawProviderValueByNameWithMessagePump`. If (c), EV-E2/E3 become deletes.
- **Verify.** RED assertion first: prove the fallback currently never fires (trace counter), then
  fix and watch it fire on an induced UIA-miss.

### EV-E2 [LOW] Raw matcher drops the visibility contract (F#64) — *with EV-E1*
`Commands.SelfTest.Settings.cpp:3980-4004` matches control type + name only; primary path
requires `UIA_IsOffscreenPropertyId == FALSE` (`:3899`). Add an `IsOffscreen` read (VT_BOOL
sibling of `TryReadRawProviderLongProperty`).

### EV-E3 [LOW] `InvokeVisibleDescendantByName` can double-fire (F#25) — *with EV-E1*
`Commands.SelfTest.Settings.cpp:5591`. UIA Invoke on a window-closing button commonly reports
failure after the action ran; the `|| raw invoke` then fires it again. Only fall back when the
element/pattern **lookup** failed, not when `Invoke()` errored — or verify the expected
post-condition didn't occur before re-invoking.

### EV-E4 [MEDIUM] OK/Cancel keyboard backdoor silently removes UIA coverage (F#35)
`Commands.SelfTest.CompareOptions.cpp:1494`. `invokeButtonOrActivateTarget`'s second alternative
(debug focus hook + synthetic Space) involves no UIA; an InvokePattern regression now passes
untraced. Trace loudly when the fallback path is taken, or keep a separate non-fatal UIA-only
Require.

### EV-E5 [MEDIUM] `setAndWaitForEditValue` retry loop can never retry (F#44)
`Commands.SelfTest.CompareOptions.cpp:1394` (nested 3s waits exceed the 3s outer deadline; a
failing set burns ~9.5s ×4 calls). Short per-attempt timeout (250–500ms) inside the loop, delete
the duplicated second wait, keep one final full-length wait.

### EV-E6 [LOW] Illusory timeout in `CollectVisibleCompareOptionsEditDiagnostics` (F#37)
`Commands.SelfTest.CompareOptions.cpp:184` (`request_stop()` targets a lambda that ignores its
token; outer loop blocks until worker finishes). Drop the stop/timedOut machinery (single
non-cancellable pass) or mirror `RunUiaActionWithMessagePump`'s detach-on-timeout.

### EV-E7 [MEDIUM] worker_shutdown test scrapes product source text (F#36)
`CompareDirectoriesEngine.SelfTest.Cases.RuntimeAndRemote.cpp:1775`. Keep the behavioral half
(create session → observe worker → reset); drop the source-string/member-declaration-order
assertions; hoist ONE shared whitespace-insensitive Contains helper to SelfTest common (this diff
added three per-file `compactWhitespace` copies).

### EV-E8 [LOW] Failure-path shortcut check uses plain `std::filesystem::exists` (F#38)
`Commands.SelfTest.ShellCommands.cpp:3679`. Use
`QueryShellShortcutPathForShellCommandTest(staleLink, exists)` and require
`SUCCEEDED(hr) && !exists`, matching the positive checks this same diff migrated to `\\?\`.

### EV-E9 [LOW/simplification] Hoist `extractJsonUInt`/`stableDeviceHash`; exact counter assertions (F#58)
`Cases.Mtp.cpp:1905` (copy without the overflow guard; 8 copies total; 12 of stableDeviceHash).
Also replace `props.find(R"json("copyItemCalls":1)json")` (matches 10–19; lines 6192/6811, same
for maxConcurrentBackendCalls) with `extractJsonUInt(props, "copyItemCalls") == 1u`.

### EV-E10 [LOW/simplification] Hoist the 6×-copied ~130-line helper-lambda block (F#59)
`Cases.Mtp.cpp:5783` (new 6th copy; others at ~3840/4082/4340/4656/4931): `narrowAscii`,
`stableDeviceHash`, `getLocalAppDataPath`, `ensureDirectoryExists`, `writeUtf8File` →
anonymous-namespace free functions. ~650 lines removed; `stableDeviceHash` must stay bit-identical
to the plugin's journal-path hashing — one copy makes that checkable.

### EV-E11 [LOW] Delete unreachable `continue` after `Require` (F#60)
`Tests/DxUiTests/DxUiTests.NativeTextInput.cpp:198/239/270` (Require `std::exit(1)`s).

### EV-E12 [LOW] ViewerVLC self-fulfilling focus check (F#40)
`Tests/ViewerPETests/ViewerPETests.cpp:4648` (test SetFocuses the video window itself, then
asserts routing). Keep a non-latching diagnostic for product-driven routing + reword the latching
check to what it actually verifies ("video surface can take keyboard focus"), or record fallback
use so routing regressions stay visible.

---

## Track F — Tools / CI

### EV-F1 [HIGH — currently red] TestInventory count drift; fix stranded off-master (F#34)
`Tools/Tests/TestInventory.Tests.ps1:35` + `Specs/Testing/Testing_TestCoverage.md` (lines ~32,
~2008). Master added 5 RunCase registrations but bumped assertions by 4 — both the exact-count
and doc-vs-source checks fail deterministically; the correction exists only in orphan snapshot
`31a0b6b83` (2026-07-05, not on master).
- **Fix.** Re-derive the true count at execution time (`Get-RSTestInventory` over the current
  tree — the working tree adds more cases, so do NOT trust 234/235/238 from any snapshot), set
  the Pester expectation and the coverage doc to it, then delete/ignore the superseded snapshot
  ref. Consider making the doc the single source of the number (the doc-vs-source assertion
  already exists) so it can't fork again. **Do this first — it unbreaks the suite gate.**

### EV-F2 [LOW] Respect an explicit `-MinimumSamplesForP95` (F#27)
`Tools/Show-PerfRuns.ps1:505`. When `$PSBoundParameters.ContainsKey('MinimumSamplesForP95')`,
prefer the explicit value (or `max(explicit, budget)`) over the budget/one-shot lookups.

### EV-F3 [LOW] Warn on budget-file parse failure (F#28)
`Tools/Show-PerfRuns.ps1:474`. A JSON5-legal edit (single quotes, unquoted keys) silently zeroes
every budget minimum and `-FailOnQuality` passes with 1 sample. `Write-Warning` + latch
`$script:QualityFailure` when the file exists but fails to parse.

### EV-F4 [LOW/simplification] Extract `Invoke-RSPwsh`; add `-BudgetPath` override (F#61)
`Tools/Tests/ShowPerfRuns.Tests.ps1:146` (~40 duplicated lines; suite gates on the live repo
budget file, so routine budget retuning breaks tests with no product change).

### EV-F5 [LOW] Allowlist guard for .rc discovery (F#41)
`Tools/Tests/ResourceLocalizationContracts.Tests.ps1:104`. Add a test that fails if any `*.rc`
exists under a repo-root child not in the 7-root allowlist (excluding .build/packages/.claude),
so new top-level projects must consciously opt in.

### EV-F6 [LOW] Strengthen the secure-clear contract assertion (F#39) — *with EV-B7*
`Tests/DxUiTests/DxUiTests.TextField.cpp:1539`. Match the meaningful call tail
(e.g. `readingDirection, true)`) instead of the bare substring `"true)"`.

---

## Track G — Repo hygiene (human sign-off required)

### EV-G1 [MEDIUM] Delete the stale artifact-laden branch; reclaim ~3.5 GB (F#30)
Local branch `codex/folderview-warpdrive` tips at `46cdcc14b` — a byte-twin of merged master work
**plus 8 giant accidental artifacts** (two binary `staged-diff.patch` blobs of ~1.0 GB and
~2.1 GB, multi-million-line patch/jsonl files under `Specs/TestRuns/4cb089111a23/Continuation/`).
Clean twins (`53e7e6641`/`a9c1a998d`/`e814abbb6`) are on master; the dirty originals are pinned by
the branch and `refs/codex/snapshots/*`.
- **Fix.** (1) Re-verify at execution time that each twin pair differs ONLY in the 8 artifact
  files (`git diff 46cdcc14b 53e7e6641 --stat` etc. — this was verified 2026-07-05 but refs may
  move). (2) Delete the local branch, prune the snapshot refs reaching
  `46cdcc14b/6e6ff863b/52e3e70ec/1b996cfe2/b1756f801`, `git gc`. (3) Add
  `Specs/TestRuns/**/staged-diff.patch` and `**/perf_metrics.jsonl` (confirm exact patterns
  against `Tools` archival scripts first) to `.gitignore`.
- **Risk.** Destructive (ref deletion) and it removes the re-merge hazard — the same failure mode
  that silently reverted the TabControl fix (see Granite/CW-7 history). **Ask the human before
  executing**; do not run `gc --prune` while any worktree session has those refs checked out.

---

## Suggested sequencing

0. **EV-F1** — the inventory gate is red on master right now; nothing else can pass the suite
   cleanly until it lands.
1. **Track A core (EV-A1→A5, with EV-TA1 + EV-A9 first so the fixture can express RED tests)** —
   MTP viability; biggest user-facing breakage. Coordinate with the uncommitted Granite slice.
2. **EV-C1** (UI stall + dead watcher on first run), **EV-B4/B5** (TSF UAF), **EV-B7 + EV-F6**
   (password hygiene), **EV-D1/D2/D3** (wedges + wrong-context notify + temp race).
3. **EV-E1→E3** as one unit (fallback reachability decides E2/E3), then E4–E8; **EV-C3, EV-C5,
   EV-A6–A8, EV-D8**.
4. **Simplification batch** (EV-A12, B6, D4–D7, E9–E11, C4, C7, F2–F5, E12) — mechanical, low
   risk, big readability payoff; fine to interleave as files are already open for earlier items.
5. **EV-G1** with human confirmation; **EV-A10** (fixture extraction) whenever MTP files quiet
   down.

## Gates / done criteria

- Per-item: the cited behavior provably changed (RED-first where the fixture allows: EV-A1, A2,
  A4, C3, C5, E1), targeted selftest(s) green, no new inventory drift (EV-F1's assertion holds).
- Track A closeout: MTP replug + external-modification + re-Initialize scenarios pass on the WPD
  cache fixture AND the FakeBackend parity cases; `mtp_backend_command_worker_is_reused` still
  green.
- Plan closeout: full suite green, findings doc updated with per-item `DONE <date>` marks, this
  file moved to `Specs/Plans/Done/`.
