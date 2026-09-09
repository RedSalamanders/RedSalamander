> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# Three-Day Diff Review — Findings (2026-07-05)

**Scope:** every change from 2026-07-02 through 2026-07-05: base `275c04034` → `master` (`da6b438a0`) **plus** the uncommitted working tree (Operation Granite search/MTP work, 17 files). Code delta reviewed: ~92 files, +8,813/−1,457 lines (excluding `Specs/TestRuns` archives).

**Method:** 13 parallel area/lens reviewers (per-subsystem + cross-cutting architecture, concurrency, and merge-integrity lenses), 81 raw findings → 69 after dedup → every finding adversarially verified (bugs by two independent verifiers: correctness + reachability). 61 findings confirmed by the fleet, 3 more confirmed by manual follow-up verification (their verifiers hit a session limit), 3 refuted (see appendix). Commits reviewed include the `codex/folderview-warpdrive` merge, "Stabilize FolderView performance validation", "checkpoint test stabilization work", and the dependabot curl 8.21.0 / vcpkg 2026.06.24 bumps.

**Verdict counts:** 12 high / 11 medium / 38 low confirmed by the fleet, +1 medium +2 low from manual verification = **64 findings**.

## Dominant themes

1. **The new MTP WPD session/path cache has no coherent invalidation story** (7 findings, 4 high). Operation-phase failures never evict dead sessions (wedge after replug), cached item metadata has no TTL (permanent `ERROR_CRC` on externally-modified files, stale sizes → spurious `ERROR_PARTIAL_COPY`), cached COM objects outlive the per-command `CoInitializeEx` scope, and a watchdog-tripped worker is never recreated (re-`Initialize` bricked with `E_FAIL`). The pre-change code opened a session per operation and self-healed; the cache traded that resilience for speed without adding a recovery path.
2. **DxUi accessibility snapshot rebuilds got heavier, not gated** (3 findings). New O(selection²) synchronous UI-thread work on selection change, dozens of new ungated full-tree rebuild triggers with no coalescing (amplifying the previously flagged no-listener-gate issue), and the trigger block copy-pasted 13×.
3. **Test-stabilization code that cannot do its job** (5 findings). The new raw UIA-provider fallback in the Settings self-tests can never succeed on the worker-thread paths where it is wired, drops the visibility contract where it could run, and can double-fire invokes; a retry loop whose nested waits exceed its own deadline; a fixed test-count assertion stranded in unmerged snapshot commit `31a0b6b83`.
4. **Copy-paste duplication as the default reuse strategy** (~12 simplification findings). Overwrite orchestration ×3, focus-restore blocks ×3, prompt-close helper ×3, `extractJsonUInt` ×8, byte-identical in-memory readers ×2, plus ~130-line helper blocks recopied per test case.
5. **Repo hygiene:** stale `codex/folderview-warpdrive` branch keeps ~3.5 GB of accidentally committed blobs reachable and invites a disastrous re-merge; one real fix is stranded on an orphan snapshot commit.

---

## Bugs (28)

### 1. [HIGH][bug] New offscreen-selected-row snapshot loop is O(selection^2) (plus O(selection*rows) FindRowByStableId) and runs synchronously on the UI thread on every selection change

**Location:** `Common/DxUi/DxUi.Accessibility.cpp:1152`  |  **Found by:** dxui (+arch-lens)

The diff adds a per-selected-row loop to the grid snapshot: for each selected row it calls model->FindRowByStableId (linear scan in FindFilesWindow's model, RedSalamander/FindFilesWindow.cpp:1084-1091), then std::find_if over record.gridRows (which grows by one entry per offscreen selected row, so the scan is quadratic in selection size), then GetCellData + BuildGridCellAccessibleText string building for every visible column of every selected row. Because this same commit also triggers a full snapshot rebuild from Grid::OnSelectAll (DxUi.Grid.cpp:3447), Grid::SelectRow (3649), Grid::OnKeyDown (3269), and every handled WM_LBUTTONUP (DxUi.WindowHost.cpp:2567), pressing Ctrl+A in a Find Files results grid (Extended selection mode, FindFilesWindow.cpp:2859) with tens of thousands of results does an O(N^2) rebuild on the UI thread — hundreds of millions to billions of operations — and repeats it on every subsequent click or arrow key while the selection persists. Concrete failure: search returning 50k results, Ctrl+A → multi-second-to-minutes UI hang; each further keystroke re-hangs. This lands in a branch whose stated purpose is FolderView performance remediation. The snapshot also retains one string record per selected row, ballooning snapshot memory for large selections.

**Evidence:** const auto existingRow = std::find_if(record.gridRows.begin(), record.gridRows.end(), [rowId](const AccessibilityGridRowSnapshotRecord& row) noexcept { return row.rowId == rowId; });  ... for (const size_t columnIndex : record.gridVisibleColumns) { GridCellData cellData{}; model->GetCellData(selectedRowIndex.value(), columnIndex, cellData); ... } record.gridRows.push_back(std::move(rowRecord));  (loop over `for (const uint64_t rowId : selection)` at lines 1146-1179)

**Suggested fix:** Cap the number of materialized offscreen selected-row records (e.g., first 100), build a hash set of visible rowIds once instead of find_if per row, and/or resolve offscreen row names lazily in the GridRow provider instead of eagerly in every snapshot rebuild.


### 2. [HIGH][bug] TSF deactivation reorder leaves the text store wired to raw _host/_control pointers during Pop, creating a use-after-free window when the control was already destroyed

**Location:** `Common/DxUi/DxUi.NativeTextInput.cpp:842`  |  **Found by:** dxui

Old code called DisconnectNativeTextInputTextStore() as the FIRST step of DeactivateNativeTextInputTsf, nulling the store's _host/_control before Pop(TF_POPF_ALL) could provoke synchronous TSF callbacks. The new code deliberately moves Disconnect/Detach to AFTER Pop and after the com_ptr resets (lines 887/895), so during Pop the store still holds raw _host/_control. The host never learns synchronously when a control is destroyed — TextField::~TextField (DxUi.TextInput.cpp:925) does not notify the host, and staleness is only discovered lazily via PruneStaleInteractionState (DxUi.WindowHost.cpp:3842-3845) or OnKillFocus (3754). Failure scenario: app code replaces a panel's children destroying the focused TextField while a native session (worse: an active IME composition) exists; the next OnKillFocus/prune runs DeactivateNativeTextInputSession → DeactivateNativeTextInputTsf → Pop(TF_POPF_ALL); TSF terminates the composition, requesting a read/write lock; the store's ReadState/ApplyState then execute dynamic_cast<TextField*>(_control) and textField->SetTextAndNotify(...) (DxUi.TextStoreACP.cpp:876, 912-925) on the freed control — dynamic_cast on a dangling pointer is immediate UB; GetTextExt/GetACPFromPoint also deref _control (TextStoreACP.cpp:682/688/739). Null checks in the store cannot catch a dangling pointer. The reorder fixes the live-control teardown path but silently reverts the property that deactivation was safe against stale _nativeTextInputControl.

**Evidence:** const HRESULT hr = documentMgr->Pop(TF_POPF_ALL);  ... (later, lines 887-896) DisconnectNativeTextInputTextStore(textStoreToDisconnect.get()); ... DetachNativeTextInputTextStore(textStoreToDisconnect.get()); _nativeTextInputTsfTextStore.reset();  — versus old code whose first statement was DisconnectNativeTextInputTextStore(_nativeTextInputTsfTextStore.get());

**Suggested fix:** Before Pop, verify _nativeTextInputControl is still alive/in-tree (ControlBelongsToTree or lifetime token); if stale, Disconnect the store first (old ordering) and only use the keep-attached ordering when host and control are known-live. Better: have the store hold a std::weak_ptr<int> lifetime token and check it in every entry point that derefs _control.


### 3. [HIGH][bug] Backend command worker is never recreated after watchdog trip or Disconnect, permanently bricking a re-initialized FileSystemMtp instance with E_FAIL

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:2241`  |  **Found by:** mtp-core (+arch-lens, concurrency-lens, concurrency-lens)

The new MtpBackendCommandQueue is created only in the constructor and in AbandonBackendSessionLocked's guarded branch `if (! _disconnected && _backend)`. Both call sites of AbandonBackendSessionLocked (RunBackendCommand timeout path, line 2303-2309, and ExecuteDriveMenuCommand cleanup, line 3665-3672) set `_disconnected = true` immediately before calling it under the same _stateMutex, so the recreate branch is dead code and `_backendWorker` stays null forever after any watchdog trip or Disconnect. Initialize (line 2582) resets `_disconnected = false` (the documented re-arm path, and FolderWindow.FileSystem.cpp:5135 re-Initializes the cached pane instance when the instance context changes), after which every RunBackendCommand passes the `_disconnected` check but hits `if (! backend || ! backendWorker) return E_FAIL;` (line 2284). At the base commit, RunBackendCommand only needed `_backend` + `_deviceIoMutex`, both restored by AbandonBackendSessionLocked, so re-initialization recovered with a fresh WPD session. Concrete scenario: MTP command times out (device stall) or user picks the Disconnect drive-menu item; user navigates back to the device (new connection context) — Initialize succeeds but all enumeration/copy/read operations fail with E_FAIL until the whole instance is destroyed. Note for the fix: recreating the worker must also account for still-live MtpBackendReader objects whose IMtpBackendFileReader is bound to the abandoned backend — their commands would then run on the new queue concurrently with a quarantined stuck command on the old stream.

**Evidence:** void FileSystemMtp::AbandonBackendSessionLocked(...) noexcept
{
    _deviceIoMutex = std::make_shared<std::mutex>();
    ...
    if (! _disconnected && _backend)   // always false: both callers set _disconnected = true first
    {
        _backendWorker = std::make_shared<MtpBackendCommandQueue>(_backend, _deviceIoMutex);
    }
}
...
HRESULT FileSystemMtp::Initialize(...) { ... _disconnected = false; return S_OK; }  // line 2582, worker not recreated
...
if (! backend || ! backendWorker)
{
    return E_FAIL;   // line 2284-2287, permanent after re-Initialize
}

**Suggested fix:** In Initialize (or in AbandonBackendSessionLocked without the _disconnected guard), recreate _backendWorker when it is null and _backend is set; decide explicitly how outstanding MtpBackendReader instances bound to the abandoned backend should behave (e.g., invalidate them).


### 4. [HIGH][bug] Dead cached WPD session is never evicted because operation-phase failures bypass cache invalidation, permanently wedging the MTP filesystem after a device replug

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:1427`  |  **Found by:** mtp-device

The new session/path caches are only invalidated via FailAndMaybeInvalidateCaches(), which is wired exclusively into the RESOLVE phase (ResolvePathCached/GetOrOpenSession/FindChildByDisplayNameForBackend). Every operation-phase error is returned raw: EnumerateDirectory returns enumHr (line 1427), ReadFile returns readHr (line 1500), DeleteItem returns deleteHr (line 1641), RenameItem returns renameHr (1678), CreateFileReader/UploadFileObjectCached use RETURN_IF_FAILED. GetOrOpenSession (lines 1997-2003) hands back the cached IPortableDevice/IPortableDeviceContent without any liveness check. Failure scenario: user browses a phone (paths + session cached), unplugs and replugs it (or the WPD service hiccups). Refreshing the current folder: ResolvePathCached hits the full-path cache (zero device I/O), then EnumerateObjectItems on the dead session fails promptly (e.g. E_WPD_DEVICE_NOT_OPEN) and the error is returned WITHOUT invalidating anything. Every retry is identical, forever. The Core-side recovery (AbandonBackendSessionLocked) only triggers on watchdog TIMEOUT or the manual Disconnect menu — a prompt failure never trips it. The pre-change code (base 275c04034, old ResolvePath line 1022) opened a fresh session per operation and self-healed; this is a regression introduced by the caching work.

**Evidence:** Device.cpp 1424-1427: `const HRESULT enumHr = EnumerateObjectItems(resolved.content, resolved.objectId, items); if (FAILED(enumHr)) { return enumHr; }` — no FailAndMaybeInvalidateCaches. GetOrOpenSession 1997-2002: `if (const auto cached = _sessionsByPnpId.find(key); cached != _sessionsByPnpId.end() && DesiredAccessIsCovered(...)) { device = cached->second.device; content = cached->second.content; return S_OK; }` — cached session returned unvalidated. Core.cpp RunBackendCommand 2291-2309: backend is only abandoned when `status->cv.wait_for(...)` times out.

**Suggested fix:** Route operation-phase failures through FailAndMaybeInvalidateCaches (at minimum evict the _sessionsByPnpId entry and the pnpId's path-cache subtree on any non-whitelisted HRESULT), so the next attempt reopens a fresh session.


### 5. [HIGH][bug] Path cache never refreshes cached MtpItem metadata (no TTL, cache-hit skips live lookup), so a file modified on the device makes ReadFile fail permanently with ERROR_CRC and GetSize/GetFileSize report stale sizes

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:1497`  |  **Found by:** mtp-device

TryResolveCachedPath (line 2042 `resolved.item = cached->second.item;`) returns the MtpItem captured at first resolution and skips all device I/O; nothing ever refreshes an existing entry on success (EnumerateDirectory fetches fresh items but does not write them into _pathCache), and the cache has no expiry. ReadFile passes the cached `resolved.item.sizeBytes` as expectedSizeBytes into ReadPortableDeviceStream, which returns HRESULT_FROM_WIN32(ERROR_CRC) on any mismatch (lines 1018-1022) — and per the previous finding that error does not invalidate the cache. Failure scenario: user previews/copies a file once (item cached), the phone then modifies it (in-progress video recording, download, messaging app rewrite), user copies it again: the stream reads the new full content, size differs from the stale cached size, ReadFile returns ERROR_CRC — on every subsequent attempt, forever, even though the folder listing (live enumeration) shows the correct new size. Similarly WpdStreamBackendFileReader::GetSize (line 1145) and GetFileSize (line 1582) report the stale size; the host move flow (FolderWindow.FileOperations.State.cpp:8329) compares destination size against reader GetSize and raises spurious ERROR_PARTIAL_COPY. Base code resolved fresh metadata per operation, so mismatches were only transient.

**Evidence:** Device.cpp 1497: `ReadPortableDeviceStream(resolved.content, resolved.objectId, resolved.item.sizeBytes, bytes)` where resolved.item came from `resolved.item = cached->second.item;` (2042). ReadPortableDeviceStream 1018-1021: `if (bytes.size() != expectedSizeBytes) { ... return HRESULT_FROM_WIN32(ERROR_CRC); }`. StoreResolvedPath (2050-2063) is only called on cache misses; no refresh/TTL exists.

**Suggested fix:** Refresh or drop the cached item when it is about to be used for size-sensitive operations (e.g. re-fetch object properties for files before read/transfer), or add a TTL, or update _pathCache entries from every successful live enumeration; also invalidate the entry when the size check fails.


### 6. [HIGH][bug] New WPD path cache has no staleness mechanism: missing-object failures and fresh enumerations never invalidate cached entries, so externally-changed device files become permanently unopenable or mis-sized

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:2090`  |  **Found by:** arch-lens

The new _pathCache caches objectId/size/attributes per path with no TTL. Two gaps: (1) ShouldInvalidateAllCachesOnFailure deliberately returns false for ERROR_FILE_NOT_FOUND / ERROR_PATH_NOT_FOUND, and failure paths after a cache hit bypass FailAndMaybeInvalidateCaches entirely (e.g. CreateFileReader's 'RETURN_IF_FAILED(resources->GetStream(objectIdCopy...))' at Device.cpp:1544). (2) EnumerateDirectory fetches fresh items via EnumerateObjectItems but never stores or refreshes _pathCache entries, so the listing and the cache are two divergent sources of truth. MTP device content changes externally as the normal case (camera apps delete+recreate files with new objectIds): the listing shows the current file, but open/copy resolves through the stale cache entry, GetStream fails on the dead objectId, and every retry hits the same cached entry forever — the file is permanently unopenable until an unrelated 'unexpected' error class happens to clear all caches. Similarly, GetFileSize (Device.cpp:1580-1600) and the reader's GetSize silently return the stale cached sizeBytes after an in-place size change, so copy engines can mis-size transfers.

**Evidence:** case HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND):
case HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND):
...
    return false;   // never invalidates on missing-object
// CreateFileReader (1544): RETURN_IF_FAILED(resources->GetStream(objectIdCopy.c_str(), ...));  // stale objectId failure, no cache invalidation
// EnumerateDirectory: fresh EnumerateObjectItems results are never written back to _pathCache

**Suggested fix:** Invalidate the specific _pathCache entry (and descendants) when an operation on a cache-resolved object fails with a missing-object error, and refresh/replace cached child entries from EnumerateDirectory results (or add a TTL).


### 7. [HIGH][bug] New WPD path cache is never invalidated for device-side (external) changes, and operation failures on stale objectIds do not heal it, causing persistent wrong results

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:2020`  |  **Found by:** concurrency-lens (+mtp-device)

TOCTOU on device state: _pathCache maps path -> {objectId, MtpItem metadata} with no TTL and no revalidation on hit. MTP devices mutate their own content (photos taken/deleted/renamed on the phone). After an external delete+recreate of a file: (a) GetAttributes/GetBasicInformation/GetFileSize return the STALE cached attributes/size with S_OK and never touch the device (Device.cpp:1441-1449 return resolved.item from cache); (b) CreateFileReader/DeleteItem/RenameItem resolve the stale objectId and the subsequent WPD call fails — but those failures are returned directly (resources->GetStream at 1544, DeleteObjectById at 1636, RenameObjectById at 1672) without ever passing through FailAndMaybeInvalidateCaches, and ShouldInvalidateAllCachesOnFailure (2090-2114) deliberately excludes ERROR_FILE_NOT_FOUND/PATH_NOT_FOUND anyway. So copying a file that visibly exists in the (live) directory listing fails with FILE_NOT_FOUND on every retry until an unrelated session-level error or app restart clears the caches. The old per-call ResolvePath enumerated live and never had this failure mode.

**Evidence:** Device.cpp:2020-2047 TryResolveCachedPath returns cached objectId/item with no device round-trip; 2090-2114 'case HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND): ... return false;' (no invalidation on missing); 1633-1641 DeleteItem: 'const HRESULT deleteHr = DeleteObjectById(...); if (SUCCEEDED(deleteHr)) { InvalidatePathCacheSubtree(path); } return deleteHr;' — failure path leaves the stale entry.

**Suggested fix:** Invalidate the specific path-cache entry (and retry once with a fresh resolve) whenever an operation using a cached objectId fails; or treat missing-object errors on cache-hit resolves as a signal to evict and re-resolve.


### 8. [HIGH][bug] Cached WPD COM sessions and streams outlive the per-command CoInitializeEx/CoUninitialize scope on the queue worker thread

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:2401`  |  **Found by:** concurrency-lens

Every backend entry point creates a scoped ComInitialization (Device.cpp:44-65: CoInitializeEx(MTA) in ctor, CoUninitialize in dtor), so between commands the single MtpBackendCommandQueue worker thread has COM torn down. The new _sessionsByPnpId cache (IPortableDevice/IPortableDeviceContent) and WpdStreamBackendFileReader (_content/_stream, Device.cpp:1124-1233) hold WPD COM interface pointers ACROSS those scopes. If the worker's CoUninitialize releases the process's last MTA reference, COM tears down the apartment and can unload in-proc servers while the cached interface pointers persist; the next command re-initializes COM and calls into the stale objects -> RPC_E_DISCONNECTED or access violation in an unloaded module. Additionally, WpdStreamBackendFileReader::Seek/Read (1170-1215) are executed via MtpBackendReader -> RunBackendCommand lambdas that contain no ComInitialization at all, so IStream::Read/Seek on the WPD stream run on a thread where COM may not be initialized — a COM rules violation the old code (per-command sessions fully created and released inside one COM scope; pure in-memory reader) never had. Racing actors: the queue worker's COM teardown between command N and N+1 vs the cached WPD proxies used by command N+1.

**Evidence:** Device.cpp:49-58 'ComInitialization() : hr(CoInitializeEx(nullptr, COINIT_MULTITHREADED)) ... ~ComInitialization() { if (uninitialize) CoUninitialize(); }' per call; Device.cpp:2401-2402 '_sessionsByPnpId; _pathCache;' hold wil::com_ptr<IPortableDevice>/<IPortableDeviceContent> across calls; Device.cpp:1200-1215 WpdStreamBackendFileReader::Read calls _stream->Read with no ComInitialization member or local.

**Suggested fix:** Keep COM initialized for the lifetime of the queue worker thread (CoInitializeEx once in WorkerMain, CoUninitialize at exit) and drop the per-call scopes, so cached WPD objects never cross a COM teardown; add ComInitialization to the stream-reader paths.


### 9. [HIGH][bug] The new raw UIA-provider fallback can never succeed on the worker-thread paths where it is wired, so the stabilization it was added for is a silent no-op.

**Location:** `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp:5314`  |  **Found by:** selftests-a

All 'WithMessagePump' helpers run their action on a fresh std::jthread (RunUiaActionWithMessagePump, Settings.cpp:5289-5293) while the self-test/UI thread pumps. The new raw-provider fallbacks call RedSalamander::DxUi::CreateWindowHostAccessibilityProvider(), which returns nullptr off the window thread: it requires target->ResolveHost() != nullptr (DxUi.Accessibility.cpp:7108-7118), and ResolveHost() explicitly rejects callers whose thread id differs from GetWindowThreadProcessId(hwnd) (DxUi.Accessibility.cpp:353-357). The Commands self-test runs in the main window's wndproc (RedSalamander.cpp:11098 -> RunCommandsSelfTestAndRequestShutdown), so every window under test lives on the main thread and the worker thread always fails the gate. Consequences: (1) SetWindowHostRawProviderValueByNameWithMessagePump (Settings.cpp:5314-5327) unconditionally returns false — it exists only to run the raw path on a worker; (2) the fallbacks inside SetVisibleDescendantValue (5538), InvokeVisibleDescendantByName (5582/5588/5591), CollectVisibleDescendantValuePatternStateByName (5190) and CollectVisibleDescendantNamedElementState (4960) are inert whenever those helpers execute under a pump wrapper — which is exactly how CompareOptions uses them (setNamedEditValue, CompareOptions.cpp:1360-1364 tries the always-false raw set first, then falls back to the same flaky UIA path as before). When the UIA client path flakes (the motivating failure), behavior is unchanged from before the fix, plus every call now spawns a useless worker thread and EnumChildWindows walk. Everything except provider *creation* was designed for cross-thread use (reads are snapshot-based, actions marshal via SendMessageTimeoutW), so the creation-time thread gate is what defeats the design.

**Evidence:** Settings.cpp:5324-5326: return RunUiaActionWithMessagePump(L"raw ValuePattern SetValue", label, [hwnd, ...]() noexcept { return SetWindowHostRawProviderValueByName(hwnd, ...); });  —  DxUi.Accessibility.cpp:353-357 (ResolveHost): if (windowThreadId != 0u && windowThreadId != GetCurrentThreadId()) { return nullptr; }  —  DxUi.Accessibility.cpp:7110-7117 (CreateWindowHostAccessibilityProvider): if (! target || target->ResolveHost() == nullptr) { ... return nullptr; }

**Suggested fix:** Either create the provider on the window thread (e.g., marshal creation via SendMessage to a test hook, or relax the ResolveHost gate for the ENABLE_TESTS factory since reads are snapshot-based and actions already marshal), or drop the raw fallback from the worker-side helpers and delete SetWindowHostRawProviderValueByNameWithMessagePump.


### 10. [HIGH][bug] New 2s blocking readiness wait in SettingsHotReload::Start stalls the UI thread and permanently disables hot reload when the settings directory does not exist yet (first run), regressing the previously self-healing watcher

**Location:** `RedSalamander/SettingsHotReload.cpp:518`  |  **Found by:** compare-prefs (+arch-lens, concurrency-lens)

Start() now blocks the calling thread with WaitForSingleObject(readyHandle, 2000) and calls Stop() (killing the watcher for the whole session) if the ready event is not signaled. The watcher thread only signals readyEvent AFTER FindFirstChangeNotificationW succeeds (SettingsHotReload.cpp:88-103); on failure it retries every 1000ms and never signals. The settings directory is only created on first save (WriteFileBytesAtomic -> wil::CreateDirectoryDeepNoThrow, Common/Common/SettingsStore.cpp:384); LoadSettingsWithRecoveryInfo returns S_FALSE with defaults for a missing file without creating anything (SettingsStore.cpp:5218-5238), and SaveAppSettings runs at shutdown. Start() is called from OnMainWindowCreate (RedSalamander/RedSalamander.cpp:8975). Failure scenario: fresh install with no %LOCALAPPDATA%\<company>\<settings> directory -> FindFirstChangeNotificationW fails with ERROR_FILE_NOT_FOUND -> main-window creation stalls ~2 seconds, Start returns HRESULT_FROM_WIN32(WAIT_TIMEOUT), Stop() joins and tears down the watcher, and settings hot reload stays dead for the entire session. Before this change Start returned S_OK immediately and the retry loop began watching as soon as the directory appeared (e.g. after the first Apply in Preferences). Same regression applies to any transient startup failure that previously self-healed.

**Evidence:** const DWORD waitResult = WaitForSingleObject(readyHandle, 2000);
if (waitResult != WAIT_OBJECT_0)
{
    Stop();
    if (waitResult == WAIT_TIMEOUT)
    {
        return HRESULT_FROM_WIN32(WAIT_TIMEOUT);
    }
... (SettingsHotReload.cpp:518-529)  -- ready event only set after success: HANDLE raw = FindFirstChangeNotificationW(...); if (raw == nullptr || raw == INVALID_HANDLE_VALUE) { ... WaitForSingleObject(stopEventHandle, 1000) ... continue; } changeNotification.reset(raw); if (readyEventHandle) { SetEvent(readyEventHandle); } (SettingsHotReload.cpp:88-103)

**Suggested fix:** Signal readyEvent as soon as the thread has started (before the create-notification retry loop), treating 'thread alive and retrying' as success; or keep the handshake but on WAIT_TIMEOUT leave the watcher running instead of calling Stop(); or ensure the settings directory exists (CreateDirectoryDeep) before starting the watcher.


### 11. [MEDIUM][bug] secureCacheText flag does not cover the cache-replacement path — previous cached (potentially revealed-password) text is overwritten/deallocated without secure wipe

**Location:** `Common/DxUi/DxUi.SingleLineTextEditing.cpp:648`  |  **Found by:** dxui

The diff threads secureCacheText through GetOrCreateSingleLineTextLayout so TextField's layout cache is wiped securely — but only on the three failure/clear paths. On the common path (text changed, new layout created), `cache->text = std::wstring(text);` plainly assigns over the old string: if the new text is longer, the old buffer is freed un-wiped; if shorter, trailing characters of the previous text remain in the buffer past the new length. TextField passes real text into this cache whenever the field is unmasked or a password is revealed (GetDisplayText returns _text for PasswordRevealMode::Visible, DxUi.TextInput.cpp:3107-3110), so while typing/erasing in a revealed password field every keystroke leaks the prior text prefix into unwiped heap memory — defeating the purpose of the flag this diff adds (and of the broader password-wipe remediation).

**Evidence:** if (cache)
{
    cache->layout = layout;
    cache->text   = std::wstring(text);   // no SecureClear of previous contents when secureCacheText is true

**Suggested fix:** When secureCacheText is true, SecureWipe::SecureClear(cache->text) before assigning the new value (and assign via a buffer-exact copy so no stale tail survives capacity reuse).


### 12. [MEDIUM][bug] Global overwrite-journal absent-cache races across FileSystemMtp instances sharing a device identity, letting read-path replay skip a live journal

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:1392`  |  **Found by:** mtp-core

g_absentOverwriteJournalIdentities is process-global while journal file operations are only serialized per instance (each instance has its own command-queue worker). Two panes on the same device share one journal file and one cache key. Race: instance B runs RecordOverwriteJournalIntent — it invalidates the cache at entry (line 1134) and then writes the journal (line 1160) WITHOUT re-invalidating after the write (unlike WriteOverwriteJournalEntry, which invalidates after a successful write at line 1086-1090). Interleaving: B invalidates -> A finishes its own commit, DeleteFileW succeeds in ClearOverwriteJournalIntent (line 1181) -> B writes its journal -> A executes MarkOverwriteJournalAbsent (line 1192). The cache now says 'absent' while B's journal exists on disk. If B's overwrite then fails in the rename-temp window (destination already deleted, temp retained, journal kept via clearJournal=false), every subsequent read/metadata command skips replay at line 1392 (`useAbsentCache` is true for reads), so the destination file stays missing in enumerations until some mutating command (which passes cacheOverwriteJournalAbsence=false) or a process restart finally replays. The same stale-absent effect occurs deterministically if two instances derive case-differing identities (cache key is case-sensitive, journal path hash uses towlower). The equivalent interleave also exists between RecordOverwriteJournalIntent's early invalidate and another instance's replay probe marking absent (line 1409-1412).

**Evidence:** HRESULT ReplayOverwriteJournal(...)
{
    if (context.useAbsentCache && IsOverwriteJournalAbsentCached(context.deviceIdentity))
    {
        return S_OK;   // skips probing a journal that actually exists
    }
...
// RecordOverwriteJournalIntent:
    InvalidateOverwriteJournalAbsent(context.deviceIdentity);  // line 1134, before the write
    ...
    hr = WriteUtf8FileAtomic(journalPath, json);               // line 1160, no invalidate after success
...
// ClearOverwriteJournalIntent:
    token.recorded = false;
    MarkOverwriteJournalAbsent(token.deviceIdentity);          // line 1192, can overwrite B's invalidation

**Suggested fix:** Re-invalidate after the successful write in RecordOverwriteJournalIntent, and make mark-absent conditional on no intervening invalidation (e.g., per-identity generation counter under g_overwriteJournalAbsentMutex); normalize the cache key with the same towlower folding as StableDeviceHash.


### 13. [MEDIUM][bug] Cached WPD COM objects are used and released outside any COM initialization scope: reader IStream calls run on the queue worker without CoInitializeEx, and cached sessions can be released after per-call CoUninitialize tore down the MTA

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:1209`  |  **Found by:** mtp-device

Every IMtpBackend method carefully scopes COM via ComInitialization (CoInitializeEx(MTA)/CoUninitialize per call, lines 44-65), but the new caching work retains WPD COM pointers beyond those scopes: _sessionsByPnpId (line 2401) holds IPortableDevice/IPortableDeviceContent across calls, and WpdStreamBackendFileReader (1124-1229) holds IPortableDeviceContent + IStream. MtpBackendReader::GetSize/Seek/Read (Core.cpp 2111-2180) submit lambdas that call these IStream methods directly on the MtpBackendCommandQueue worker with NO ComInitialization at all — the only WPD call path in the plugin without the guard. If the per-call ComInitialization is the process's only MTA participant, the CoUninitialize at the end of CreateFileReader performs MTA rundown; the cached in-proc WPD objects created in that apartment are then invalid, and the next reader->Read / next backend call using a cached session is use-after-apartment-teardown (RPC_E_DISCONNECTED, E_WPD_DEVICE_NOT_OPEN, or a crash inside PortableDeviceApi). The final Release of cached sessions also happens on whatever thread drops the last shared_ptr (e.g. the quarantined worker thread exiting, after its COM scope ended). CLSID_PortableDeviceFTM makes cross-thread use legal, but only while the MTA stays alive and COM is initialized on the calling thread.

**Evidence:** Device.cpp 1207-1209 (Read) and 1177-1179 (Seek): `_stream->Read(buffer.data(), request, &bytesRead);` with no ComInitialization in scope; Core.cpp 2176-2180: lambda `return backendReader->Read(std::span<std::byte>(*scratch), bytesToRead, *result);` runs in WorkerMain (Core.cpp 1973-2031) which never initializes COM; Device.cpp 49: `ComInitialization() noexcept : hr(CoInitializeEx(nullptr, COINIT_MULTITHREADED)), uninitialize(SUCCEEDED(hr))` with `~ComInitialization() { ... CoUninitialize(); }`.

**Suggested fix:** Initialize COM once for the lifetime of the MtpBackendCommandQueue worker thread (all backend and reader calls already funnel there), or have WpdMtpBackend hold a CoIncrementMTAUsage cookie for its lifetime; keep per-call init only as a fallback.


### 14. [MEDIUM][bug] Shortcut save temp-file name reuse race can install a shortcut with the wrong target or fail spuriously under concurrent pastes

**Location:** `RedSalamander/FolderViewInternal.h:1215`  |  **Found by:** folderview-fileops

The diff removed the direct persist->Save fast path for paths < MAX_PATH, so EVERY shortcut save now goes through CreateShellPersistTempPath, which calls GetTempFileNameW (reserving the name by creating the file) and then immediately DeleteFileW's it (FolderViewInternal.h:1197) before persist->Save writes it later. Between the delete and the MoveFileExW, the name is free for reuse: GetTempFileNameW with uUnique=0 seeds from the current system time, so a concurrent saver (paste-shortcut in the other pane — each FolderView has its own _pasteShortcutInFlight, so cross-pane saves run concurrently; or a second RedSalamander instance sharing %TEMP% and the same 'rsl' prefix) can re-create the identical rslXXXX.tmp name. Both workers then persist->Save to the SAME temp path; the first MoveFileExW installs whichever content was written last (a .lnk pointing at the other paste's target) into its destination slot and deletes the temp, and the second worker's VerifyShellShortcutExactPath/MoveFileExW fails with ERROR_FILE_NOT_FOUND. This partially reintroduces, cross-pane, the same wrong-shortcut-content class of race that this very change (GR-2) was written to eliminate.

**Evidence:** FolderViewInternal.h:1191-1202: `if (GetTempFileNameW(tempDirectory.data(), L"rsl", 0, tempFile.data()) == 0) ... tempPath = tempFile.data(); if (DeleteFileW(tempPath.c_str()) == 0)` — the name is reserved then immediately freed. FolderViewInternal.h:1214-1229: `HRESULT hr = CreateShellPersistTempPath(tempPath); ... hr = persist->Save(tempPath.c_str(), TRUE);` — Save happens long after the delete. The removed fast path (old code: `if (linkPathText.size() < MAX_PATH) { persist->Save(linkPathText.c_str(), TRUE); ... }`) previously kept short-path saves off this machinery.

**Suggested fix:** Do not delete the reserved temp file before Save (IPersistFile::Save with STGM_CREATE overwrites the empty placeholder), or generate the temp name with a GUID; ideally create the temp file in the destination folder so the final MoveFileEx is an atomic same-volume rename instead of a cross-volume copy from %TEMP%.


### 15. [MEDIUM][bug] mtp_wpd_session_and_path_cache_reuse lacks the ERROR_PROC_NOT_FOUND skip guard and hard-fails on release plugin builds

**Location:** `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp:2220`  |  **Found by:** selftests-b

Every other MTP selftest case in this file checks `createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND)` and returns state.Skip(L"MTP fake backend export is available only in debug plugin builds.") because both selftest factory exports (RedSalamanderMtpCreateForSelfTest and RedSalamanderMtpCreateWpdCacheForSelfTest) are compiled only under `#ifdef _DEBUG` (Plugins/FileSystemMtp/Factory.cpp:119-181). The new WPD session/path-cache case goes straight to state.Require, so when the suite runs against a release-built FileSystemMtp plugin, GetProcAddress fails, TryCreateWpdCacheMtpFileSystemInstance returns ERROR_PROC_NOT_FOUND (CompareDirectoriesEngine.SelfTest.cpp:1091-1094), and this single case FAILS the suite while all 40+ sibling cases skip gracefully.

**Evidence:** const HRESULT createHr = TryCreateWpdCacheMtpFileSystemInstance("{}", R"json({"readOnly":true})json", L"/", created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP WPD cache: create selftest instance failed. hr=0x{:08X}", ...));  // no ERROR_PROC_NOT_FOUND -> Skip branch, unlike e.g. line 2105-2108 in the adjacent case

**Suggested fix:** Add the same guard as sibling cases: `if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND)) { return state.Skip(L"MTP WPD-cache selftest export is available only in debug plugin builds."); }` before the Require.


### 16. [LOW][bug] ConnectClientPipe's missing-pipe retry loop has no cancel/stop hook, and the new env override stretches the uncancellable wait to up to 30s per connect

**Location:** `Common/SearchServiceBroker.cpp:1519`  |  **Found by:** search

Previously the missing-pipe retry window was hard-capped at 250ms, so the Sleep(10) polling loop's lack of any cancellation check was negligible. The diff makes the window configurable up to kMaxMissingPipeRetryWindowMs = 30'000 ms via REDSALAMANDER_SEARCH_SERVICE_CLIENT_MISSING_PIPE_RETRY_MS (the self-test harness sets 5000ms in CompareDirectoriesEngine.SelfTest.cpp:4428). SearchServiceTree calls GetStatus and Query, each of which calls ConnectClientPipe, and ClientIoContext's cancelCheck is only consulted after the pipe is connected. Failure scenario: with the override active and the service pipe absent, a user/test cancel of the search blocks inside this loop for up to the full window per connect attempt (up to 2x per search), which directly lengthens the cancellation-latency tail the test-stabilization work is trying to reduce.

**Evidence:** if (waitError == ERROR_FILE_NOT_FOUND)
{
    if (::GetTickCount64() >= missingPipeDeadline)
    {
        return HRESULT_FROM_WIN32(waitError);
    }

    ::Sleep(10u);
    continue;
}   // lines 1517-1526 — no cancelCheck/stop event in the loop

**Suggested fix:** Thread the caller's CancelCheckFn (or a stop event) into ConnectClientPipe and test it each iteration of the missing-pipe retry loop, returning ERROR_CANCELLED promptly.


### 17. [LOW][bug] Last-instance roots cleanup uses decrement-then-lock, so a concurrently constructed instance can have the shared static _roots wiped underneath it

**Location:** `Plugins/FileSystemDummy/FileSystemDummy.cpp:2896`  |  **Found by:** compare-prefs

The destructor decides it is the last instance via _liveInstanceCount.fetch_sub(1)==1 BEFORE acquiring _mutex. Between the fetch_sub and the scoped_lock, another thread can construct a new FileSystemDummy (fetch_add 0->1, e.g. final COM Release of the old instance dropping on a worker thread while the UI thread creates a new plugin instance) and start mutating the shared inline-static _roots; the outgoing destructor then locks _mutex and ClearRootsIteratively() erases the live instance's tree, silently discarding any files/directories the dummy FS created since (deterministic regeneration restores only unmutated content). This is a residual race in the otherwise-correct fix (the old code cleared _roots on every destruction, which was strictly worse), but the fix is one line: serialize the count check under the same mutex that guards _roots.

**Evidence:** _liveInstanceCount.fetch_add(1u, std::memory_order_acq_rel); (FileSystemDummy.cpp:2812, no lock) ... if (_liveInstanceCount.fetch_sub(1u, std::memory_order_acq_rel) == 1u)
{
    std::scoped_lock lock(_mutex);
    ClearRootsIteratively();
} (FileSystemDummy.cpp:2896-2903)

**Suggested fix:** Take std::scoped_lock lock(_mutex) first in both constructor (around the increment) and destructor (around decrement + ClearRootsIteratively) so the last-instance check and the roots clear are atomic with respect to new-instance construction.


### 18. [LOW][bug] New RenameItem overwrite path never checks the source's attributes, so renaming a directory onto an existing file diverges: fake backend replaces the file with the directory while WPD hardware fails

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:3352`  |  **Found by:** mtp-core

The new allow-overwrite RenameItem lambda validates the destination (not-directory, PUID policy, source!=dest) but never fetches the SOURCE attributes before entering CommitDeviceSourceOverwriteWithTempSwap. With a directory source: on the real WPD backend, backend.CopyItem(sourceDir, guidTempLeaf, false) hits the non-leaf-preserving directory branch (FileSystemMtp.Device.cpp:1723-1726) and returns ERROR_NOT_SUPPORTED, so the operation fails midway (safely, after journal intent + temp cleanup). On the fake backend, CopyItemLocked happily deep-copies the directory tree to the temp name, the temp-swap deletes the destination file, renames the directory into place, and deletes the source — the rename-with-overwrite of a directory over a file SUCCEEDS. Fake-backed selftests can therefore certify a rename-overwrite behavior real devices reject, which is exactly the parity class this Granite work was meant to eliminate (the spec's new MoveItem-parity paragraph covers backend MoveItem but not this front-door path). MoveItem/CopyItem overwrite lambdas share the same gap, but those predate this diff; RenameItem's is new code.

**Evidence:** if (SUCCEEDED(commandHr) && destinationExisted && (destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
{
    commandHr = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);   // destination checked...
}
...
if (SUCCEEDED(commandHr) && destinationExisted)
{
    commandHr = CommitDeviceSourceOverwriteWithTempSwap(
        backend, commandSource, commandDest, true, verifyLevel, journalContext, tempPuidMissing, tempPuidPresent);   // ...source attributes never checked
}

**Suggested fix:** In the overwrite branch, GetAttributes(commandSource) and fail fast with ERROR_ACCESS_DENIED (or ERROR_NOT_SUPPORTED) when the source is a directory, matching what WPD hardware would ultimately do; apply the same guard to the CopyItem/MoveItem overwrite lambdas for consistency.


### 19. [LOW][bug] noexcept FileSystemMtp constructor now starts a std::jthread via make_shared; thread-creation failure or OOM calls std::terminate instead of failing plugin creation

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:2212`  |  **Found by:** mtp-core

Both FileSystemMtp constructors are declared noexcept (FileSystemMtp.h:90-91) and now execute `std::make_shared<MtpBackendCommandQueue>(...)`, whose constructor starts a std::jthread. std::thread/jthread construction throws std::system_error under thread exhaustion and make_shared throws std::bad_alloc; either escapes the noexcept boundary and terminates the whole application. The surrounding factory code (Factory.cpp CreateInstanceMtp, RedSalamanderMtpCreateForSelfTest) deliberately uses new(std::nothrow) and returns E_OUTOFMEMORY, so graceful resource-failure handling is an explicit goal of this module. At the base commit the constructor allocated nothing that spawns threads (threads were created per command inside RunBackendCommand, where failure was at least scoped to one command). Every instance also now owns a dedicated OS thread for its whole lifetime even if no backend call is ever made.

**Evidence:** FileSystemMtp::FileSystemMtp(IHost* host, std::unique_ptr<IMtpBackend> backend) noexcept
{
    ...
    _backendWorker = std::make_shared<MtpBackendCommandQueue>(_backend, _deviceIoMutex);  // jthread starts here; ctor is noexcept

**Suggested fix:** Wrap queue creation in try/catch (leave _backendWorker null and let RunBackendCommand surface a failure), or create the worker lazily on first Submit where an HRESULT can be returned.


### 20. [LOW][bug] Selftest WPD-cache backend returns S_OK with null device/content sessions, but DeleteItem/RenameItem/CopyItem/MoveItem lack the selftest guard and will null-deref the content pointer

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:1890`  |  **Found by:** mtp-device

OpenDeviceSessionForBackend in selftest mode resets device/content to null and returns S_OK (lines 1885-1892), so ResolvePathCached succeeds with resolved.content == nullptr. WriteFile and CreateDirectory are guarded with `if (_selfTestWpdCacheBackend) return E_NOTIMPL;` (2283-2286, 2336-2339), but the other four mutators are not: DeleteItem reaches DeleteObjectById(resolved.content, ...) (1636) which calls `content->Delete(...)` on a null wil::com_ptr — undefined behavior/crash; RenameObjectById (1672), CopyOrMoveNativeObject (1711/1761) and ReadPortableDeviceStream via CopyOrMoveFileByTransferCached (2372) have the same problem. Reachable in debug selftests because Core's readOnly gate is configurable: SetConfiguration honors `"readOnly": false` (FileSystemMtp.Core.cpp:2481), after which mutation calls flow through to the backend. A selftest exercising mutation paths against the wpd-selftest fixture crashes the process instead of getting E_NOTIMPL.

**Evidence:** Device.cpp 1885-1892: `if (_selfTestWpdCacheBackend) { ... device.reset(); content.reset(); return S_OK; }`; 1636: `const HRESULT deleteHr = DeleteObjectById(resolved.content, resolved.objectId, recursive);` with no selftest guard, versus 2283-2286: `if (_selfTestWpdCacheBackend) { return E_NOTIMPL; }` in UploadFileObjectCached.

**Suggested fix:** Add the same `if (_selfTestWpdCacheBackend) return E_NOTIMPL;` guard to DeleteItem, RenameItem, CopyItem and MoveItem (or null-check content in the ById helpers).


### 21. [LOW][bug] _backendThreadIdsOverflow is read in GetItemProperties without the _threadStatsMutex that guards its writes

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp:868`  |  **Found by:** mtp-core

RecordBackendThread writes _backendThreadIdsOverflow under _threadStatsMutex (line 905), but GetItemProperties reads it (line 868) holding only _mutex. The companion counter is correctly read through BackendThreadIdsObserved() under _threadStatsMutex, so the omission looks accidental. A quarantined worker (post watchdog trip) executing a stuck fake-backend call can run RecordBackendThread concurrently with GetItemProperties on a replacement worker, making this a formal data race (UB) on the exact instrumentation field the new mtp_backend_command_worker_is_reused contract relies on.

**Evidence:** _backendThreadIdsOverflow ? "true" : "false",   // read without _threadStatsMutex
...
void RecordBackendThread() noexcept
{
    std::lock_guard lock(_threadStatsMutex);
    ...
    _backendThreadIdsOverflow = true;               // write under _threadStatsMutex

**Suggested fix:** Add a BackendThreadIdsOverflow() accessor that takes _threadStatsMutex (or return both values from one locked accessor) and use it in GetItemProperties.


### 22. [LOW][bug] _pasteShortcutInFlight is never reset if the worker's completion message is not delivered, permanently wedging paste-shortcut for that view while silently reporting success

**Location:** `RedSalamander/FolderView.FileOps.cpp:990`  |  **Found by:** folderview-fileops

StartPasteShortcutWork sets `_pasteShortcutInFlight = true` after a successful TrySubmitThreadpoolCallback, and the ONLY reset is at the end of OnPasteShortcutComplete (line 1048), which runs only when the worker's PostMessagePayload(owned->hwnd, kFolderViewPasteShortcutComplete, ...) at line 979 succeeds and the message is dispatched. The worker deliberately ignores the post result (`static_cast<void>`). If PostMessageW fails — message queue at its 10,000 limit, or the hwnd already registered closed in the payload registry — the payload is deleted by PostMessagePayload but the flag stays true forever. Every subsequent PasteShortcutFromClipboard then takes the `_pendingPasteShortcutRequests.push_back(...); return true;` branch (line 874-885): the user sees the command 'succeed' but no shortcut is ever created again in that pane until restart. Before this change a lost completion only cost the folder refresh; the new serialization turns it into a permanent silent wedge.

**Evidence:** Line 979: `static_cast<void>(PostMessagePayload(owned->hwnd, WndMsg::kFolderViewPasteShortcutComplete, 0, std::move(payload)));` — result ignored. Line 990: `_pasteShortcutInFlight = true;`. Line 1048 (sole reset): `_pasteShortcutInFlight = false; StartNextPasteShortcutRequest();` inside OnPasteShortcutComplete, reachable only via that posted message (FolderView.cpp:714-722).

**Suggested fix:** Have the worker retry/flag post failure (e.g., set an atomic in a shared state block checked by a UI-side watchdog), or on the next PasteShortcutFromClipboard reset a stale in-flight flag after a timeout; at minimum log the post failure instead of discarding it.


### 23. [LOW][bug] _pasteShortcutInFlight is never cleared if the worker's completion post fails, permanently stalling all future paste-shortcut requests in that pane

**Location:** `RedSalamander/FolderView.FileOps.cpp:1048`  |  **Found by:** concurrency-lens

The new serialized paste-shortcut state machine clears _pasteShortcutInFlight only in OnPasteShortcutComplete (line 1048), which requires the threadpool worker's PostMessagePayload at line 979 to succeed and be pumped. PostMessagePayload can fail (message queue full at 10k messages, or the payload registry rejecting a closing hwnd); the payload is safely freed but no completion ever arrives, so the UI-thread flag stays true forever. Every subsequent PasteShortcutFromClipboard then appends to _pendingPasteShortcutRequests (line 874-885) and reports success while nothing ever runs — silent permanent breakage until the FolderView is destroyed. Racing threads: threadpool worker (failed post) vs UI thread (state machine waiting for a message that never comes). The pre-diff code had no cross-request state, so a lost completion only lost one refresh.

**Evidence:** FolderView.FileOps.cpp:979 'static_cast<void>(PostMessagePayload(owned->hwnd, WndMsg::kFolderViewPasteShortcutComplete, 0, std::move(payload)));' (result discarded, no fallback), 990 '_pasteShortcutInFlight = true;', 1048 '_pasteShortcutInFlight = false;' only reachable via the message handler.

**Suggested fix:** On PostMessagePayload failure, retry or post a sentinel; alternatively add a watchdog/timeout that clears the in-flight flag, or clear it in WM_NCDESTROY payload draining.


### 24. [LOW][bug] Completion notifies DirectoryInfoCache with the pane's CURRENT _fileSystem, which may be a different plugin than the one the shortcuts were created on

**Location:** `RedSalamander/FolderView.FileOps.cpp:1032`  |  **Found by:** concurrency-lens

The diff intentionally moved NotifyFolderContentsChanged outside the hasCurrentTarget/generation guard so navigating away still refreshes the cache. But it passes _fileSystem.get() — the pane's filesystem at completion time. Paste-shortcut only runs on the builtin filesystem (checked at request time, line 847), yet if the UI thread navigates the pane to another plugin (e.g. MTP) while the threadpool worker is still creating shortcuts, the completion notifies the (mtpFileSystem, localTargetFolder) context: DirectoryInfoCache is keyed per provider context (DirectoryInfoCache.h:139-152), so the wrong context gets poked and the builtin context — e.g. the other pane showing the same local folder without an active watcher — never learns the folder changed and keeps stale counts/listing.

**Evidence:** FolderView.FileOps.cpp:1032-1035: 'if (_fileSystem && ! result.createdLinks.empty()) { DirectoryInfoCache::GetInstance().NotifyFolderContentsChanged(_fileSystem.get(), result.targetFolder); }' with no check that _fileSystem is still the builtin filesystem the request ran against.

**Suggested fix:** Capture the IFileSystem (or its provider identity) in PasteShortcutRequest/Result and notify that context, or gate the notify on the plugin still being the builtin filesystem.


### 25. [LOW][bug] InvokeVisibleDescendantByName can double-fire the invoked action: it retries via the raw provider whenever the UIA Invoke returns a failure HRESULT, even if the action already executed.

**Location:** `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp:5591`  |  **Found by:** selftests-a

UIA client Invoke on a button that closes/destroys its host window commonly reports failure (e.g., UIA_E_TIMEOUT or a disconnect HRESULT) after the provider-side action has already run. The new '|| InvokeWindowHostRawProviderDescendantByName(...)' then invokes the same-named element a second time. Today this is latent because the raw fallback is inert on the worker threads where this helper runs (see the ResolveHost thread-gate finding), but it becomes live the moment that gate is fixed — and a double Invoke on an OK/Cancel-style button would act on the successor window/dialog state, producing confusing downstream failures.

**Evidence:** Settings.cpp:5591: return SUCCEEDED(invokePattern->Invoke()) || InvokeWindowHostRawProviderDescendantByName(hwnd, expectedControlType, expectedName);

**Suggested fix:** Only fall back to the raw invoke when the element/pattern lookup failed (the two earlier fallback branches), not when Invoke() itself returned an error; or verify the element still exists and the expected post-condition did not occur before re-invoking.


### 26. [LOW][bug] GetLastError() read after Stop() in the Start() wait-failure path reports an unrelated error code

**Location:** `RedSalamander/SettingsHotReload.cpp:527`  |  **Found by:** compare-prefs

In the WAIT_FAILED branch, Stop() is invoked before GetLastError() is read. Stop() calls ClearInvalidReloadAlert/HostClearAlert, SetEvent, and jthread::join, all of which overwrite the thread's last-error value, so the HRESULT returned to the caller (and logged at RedSalamander.cpp:8978) reflects whatever Stop() last did (often 0, mapping to E_FAIL) instead of the actual WaitForSingleObject failure cause.

**Evidence:** if (waitResult != WAIT_OBJECT_0)
{
    Stop();
    if (waitResult == WAIT_TIMEOUT)
    {
        return HRESULT_FROM_WIN32(WAIT_TIMEOUT);
    }

    const DWORD lastError = GetLastError();
    return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
} (SettingsHotReload.cpp:519-529)

**Suggested fix:** Capture const DWORD lastError = GetLastError(); immediately after WaitForSingleObject, before calling Stop().


### 27. [LOW][bug] Get-MetricQualityMinimum silently overrides an explicitly passed -MinimumSamplesForP95 for budget-listed and one-shot metrics

**Location:** `Tools/Show-PerfRuns.ps1:505`  |  **Found by:** tests-tools (+tests-tools)

The new per-metric minimum resolution consults the budget file first and the one-shot suffix heuristic second, and only falls back to $MinimumSamplesForP95. A user who explicitly runs '.\Tools\Show-PerfRuns.ps1 -Run X -Metric folder.frame.input_to_paint_us -MinimumSamplesForP95 500 -FailOnQuality' gets exit 0 with only 40 samples because the budget row (minimumSamples 40) wins over the explicit CLI value; similarly '-Metric icons.recall_avoided_count -MinimumSamplesForP95 50 -FailOnQuality' can never fail because any *_count metric resolves to 0. The explicitly requested gate is silently ignored with no indication in the output.

**Evidence:** function Get-MetricQualityMinimum([string]$MetricName) {
    $budgetMinimums = Get-FolderViewBudgetMinimumSamples
    if ($budgetMinimums.ContainsKey($MetricName)) {
        return [int]$budgetMinimums[$MetricName]
    }
    ...
    if (Test-OneShotMetric $MetricName) {
        return 0
    }
    return $MinimumSamplesForP95
}

**Suggested fix:** When $PSBoundParameters.ContainsKey('MinimumSamplesForP95'), prefer the explicit value (or take the max of it and the budget minimum) instead of unconditionally short-circuiting on the budget/one-shot lookups.


### 28. [LOW][bug] Budget-file parse failures are swallowed silently, disabling all budget-derived quality gates with no warning

**Location:** `Tools/Show-PerfRuns.ps1:474`  |  **Found by:** tests-tools

Get-FolderViewBudgetMinimumSamples parses Specs\Testing\FolderViewPerfBudgets.json5 with ConvertFrom-Json and on any exception resets the cache to an empty hashtable without emitting anything. The file's .json5 extension invites JSON5 syntax; pwsh's ConvertFrom-Json tolerates comments but not single-quoted strings or unquoted keys, so a legitimate JSON5 edit makes every budget minimum vanish: in -FolderViewPreset mode every metric then resolves to minimum 0 and '-FailOnQuality' exits 0 even for folder.frame.total_us with 1 sample. The only signal is the ShowPerfRuns Pester test 'still fails FolderViewPreset quality for low-sample distribution metrics' flipping, and only if that suite runs.

**Evidence:**     } catch {
        $script:FolderViewBudgetMinimumSamples = @{}
    }

    return $script:FolderViewBudgetMinimumSamples

**Suggested fix:** Emit a Write-Warning (and/or latch $script:QualityFailure) when the budget file exists but fails to parse, instead of silently degrading to no budget minimums.


## Architecture (5)

### 29. [HIGH][architecture] Dozens of new synchronous full-tree accessibility snapshot rebuild triggers with no UIA-listener gate and no coalescing — amplifies the previously flagged ungated-rebuild problem instead of fixing it

**Location:** `Common/DxUi/DxUi.WindowHost.cpp:2567`  |  **Found by:** dxui (+dxui)

PublishWindowHostAccessibilitySnapshot walks the entire control tree (navigation records, point-hit records, per-cell GetCellData + string building, dynamic_casts — DxUi.Accessibility.cpp:394-448) and the target is registered unconditionally at window attach (DxUi.WindowHost.cpp:1215), so RefreshWindowHostAccessibilitySnapshot (Accessibility.cpp:6970-6985) always rebuilds — there is no UiaClientsAreListening()/client-present gate. A 2026-07-04 review already flagged this ungated rebuild; this diff, instead of adding a gate, multiplies trigger sites: every handled mouse-up (WindowHost.cpp:2565-2568), Grid SetModel/SetSelectionMode/ApplyColumnLayout/NotifyDataChanged/RequestRemoveRowSelection/OnMouseUp/OnKeyDown/OnSelectAll/SelectRow (Grid.cpp:863,887,1081,1197,1224,1911,3105,3269,3447,3649), Tree NotifyDataChanged/SetSelectedItemId/RequestExpandedState/SelectVisibleIndex (Tree.cpp:340,350,408,1429), ComboBox::SetSelectedIndex, Control::SetVisible/SetEnabled/SetFocusable/SetAccessibleName/SetAccessibleHelpText (DxUi.cpp:145,164,181,441,459), Label/Button::SetText, Toggle::SetChecked. The triggers stack: a single grid click runs SelectRow's rebuild, any delegate-driven Label::SetText rebuild, and the WM_LBUTTONUP rebuild — 3 full tree walks per click; Grid::OnMouseUp's own refresh (3105) is immediately duplicated by the host-level one. The added refreshes in ExecuteSetStringValueOnWindowThread (Accessibility.cpp:6542/6550) are fully redundant since SetTextAndNotify already refreshes via TextField::SetText (TextInput.cpp:978) and ComboBox::NotifyTextChanged (ComboBox.cpp:2288). This is exactly the workload the WarpDrive perf branch is meant to remove.

**Evidence:** if (controlHandled && liveControl)
{
    RefreshWindowHostAccessibilitySnapshot(_hwnd, this);
}  (WindowHost.cpp:2565-2568) — and RefreshWindowHostAccessibilitySnapshot has no listener check: `if (! target || target->host.load(...) != host) { return; } PublishWindowHostAccessibilitySnapshot(*target, *host);` (Accessibility.cpp:6977-6984)

**Suggested fix:** Replace scattered eager rebuilds with a dirty flag on WindowHost that coalesces to one rebuild per frame (or rebuild lazily in CaptureAccessibilitySnapshot when dirty), and gate publication on UiaClientsAreListening() / first WM_GETOBJECT.


### 30. [MEDIUM][architecture] Stale local branch codex/folderview-warpdrive is an artifact-laden byte-twin of already-merged work, keeping ~3.5 GB of accidentally committed blobs reachable and inviting a disastrous re-merge

**Location:** `Specs/TestRuns/4cb089111a23/Continuation/2026-07-02_000500_folderview_warpdrive_pause_archive/staged-diff.patch:1`  |  **Found by:** merge-integrity

The last 3 days' history was rewritten to strip 8 giant test artifacts (two binary staged-diff.patch blobs of 1,015,971,314 and 2,074,859,459 bytes, plus staged-diff.patch text files of 540k/1.1M/2.5M lines and perf_metrics.jsonl files of ~600k lines under Specs/TestRuns/4cb089111a23/Continuation/). The clean twins (53e7e6641, a9c1a998d, e814abbb6) landed on master; the dirty originals (46cdcc14b, 6e6ff863b, 52e3e70ec, plus snapshots 1b996cfe2/b1756f801) remain reachable: 46cdcc14b is the TIP of local branch codex/folderview-warpdrive, and the others are pinned by refs/codex/snapshots/*. Verified: git diff between each twin pair shows ONLY the 8 artifact files — no source differs, so the branch contains nothing master needs. Failure scenario: (a) the branch shows as unmerged in 'git branch --no-merged master', so a routine cleanup merge would conflict-free ADD the ~3.5 GB artifacts to master (paths don't exist there); this repo already demonstrated the duplicate-merge failure mode (6e6ff863b is a second, parallel merge of the same branch pair as a9c1a998d); (b) 'git push --all' or pushing the branch uploads gigabytes; (c) git gc can never reclaim the blobs while the refs exist.

**Evidence:** git diff --name-status 53e7e6641 46cdcc14b -> 8 'A' entries, all under Specs/TestRuns/4cb089111a23/Continuation/ (stat: 'Bin 0 -> 1015971314 bytes', 'Bin 0 -> 2074859459 bytes', '573705 ++++', '630575 ++++'; '8 files changed, 6033681 insertions(+)'). git diff --name-status 6e6ff863b a9c1a998d and 52e3e70ec e814abbb6 -> the same 8 files as 'D'. git branch -a --contains 46cdcc14b -> codex/folderview-warpdrive (local only; origin has no such branch). git log -1 both twins: identical author date 2026-07-04 10:11:42, identical subject.

**Suggested fix:** Delete the local codex/folderview-warpdrive branch (its source content is fully on master via 53e7e6641/a9c1a998d) and prune the refs/codex/snapshots entries reaching 46cdcc14b/6e6ff863b/52e3e70ec/1b996cfe2/b1756f801, then gc to reclaim ~3.5 GB. Add the Specs/TestRuns artifact patterns (staged-diff.patch, perf_metrics.jsonl) to .gitignore to prevent recurrence.


### 31. [LOW][architecture] MtpBackendReader is not tied to the backend session that created it, so a reader can execute against a replaced backend's queue

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:2038`  |  **Found by:** concurrency-lens

MtpBackendReader captures a raw FileSystemMtp* owner plus a shared IMtpBackendFileReader, and every GetSize/Seek/Read submits to the owner's CURRENT _backendWorker (Core.cpp:2111/2139/2176) with no check that _backend is still the backend that created the reader. Today this is masked because any session abandonment sets _disconnected (and finding on line 2241 leaves _backendWorker null), so reader calls fail fast. But the moment the reconnect path is fixed (worker recreated in Initialize), a stale reader created before a watchdog trip would run its old backend's IStream calls on the NEW queue thread concurrently with the quarantined old worker that may still be hung inside the same WPD session — exactly the single-transport serialization the per-backend queue exists to guarantee. There is no generation/session token to reject stale readers cleanly.

**Evidence:** Core.cpp:2107-2131 (GetSize): 'const HRESULT hr = _owner->RunBackendCommand([backendReader, result](IMtpBackend&) noexcept { return backendReader->GetSize(*result); });' — the IMtpBackend& parameter is ignored; the captured backendReader belongs to whichever backend existed at CreateFileReader time.

**Suggested fix:** Store the shared_ptr<IMtpBackend> (or a session generation) in MtpBackendReader and fail with ERROR_DEVICE_NOT_CONNECTED when it no longer matches the owner's current backend.


### 32. [LOW][architecture] Hard-coded selftest fixture backend and instrumentation are embedded inside the production WpdMtpBackend class instead of behind the existing backend seam

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:1833`  |  **Found by:** mtp-device

The production backend now carries a second, fixture-driven backend inside it: SelfTestDeviceDescriptor/SelfTestItem plus `if (_selfTestWpdCacheBackend)` branches in EnumerateDevicesForBackend, OpenDeviceSessionForBackend, EnumerateObjectItemsForBackend, CreateFileReader, UploadFileObjectCached, CreateFolderObjectCached and GetItemProperties (~120 lines of fixture data and forks in production hot paths). The project already has a dedicated test double (FileSystemMtp.FakeBackend.cpp) behind the IMtpBackend seam; the divergence produces the partial coverage bug above (two mutators guarded, four not) and means GetInfo() reports liveWpd=true for the selftest backend, so Core's AbandonBackendSessionLocked (FileSystemMtp.Core.cpp:2237-2239) would swap a real device-touching WpdMtpBackend into a selftest instance on a watchdog trip. Additionally the monotonically-increasing instrumentation counters are now serialized into the production "wpd" GetItemProperties JSON (lines 1798-1812), so two identical calls return different payloads — noise for any consumer that compares or displays the properties JSON.

**Evidence:** Device.cpp 1833-1840 SelfTestDeviceDescriptor, 1867-1877/1879-1895/1897-1929 the three ForBackend forks, 1798-1812 instrumentation in the JSON for both `"wpd"` and `"wpd-selftest"`, 1369-1372 `GetInfo` returning `.liveWpd = true` unconditionally.

**Suggested fix:** Extract the caching/resolution layer to operate over a small device-ops interface and implement the selftest fixture as a separate implementation of that interface (like FakeBackend); report liveWpd=false for the selftest variant; emit instrumentation only for the selftest backend.


### 33. [LOW][architecture] A second fake MTP device (self-test fixture mode) is compiled into the production WpdMtpBackend in release builds

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:1359`  |  **Found by:** arch-lens

WpdMtpBackend now carries a '_selfTestWpdCacheBackend' mode with hard-coded fixture data (SelfTestDeviceDescriptor 'Fake Phone', SelfTestItem tree at Device.cpp:1863-1948, a canned CreateFileReader payload at 1528-1535, E_NOTIMPL write paths at 2283/2336, and a 'wpd-selftest' branch in GetItemProperties at 1800). Only the factory CreateSelfTestWpdMtpBackend is #ifdef _DEBUG (Device.cpp:2418), so all fixture code and per-call branch checks ship in release builds, and the project already has a full FakeMtpBackend for exactly this role. Test fixtures embedded in the production backend blur the backend seam the IMtpBackend interface exists to provide.

**Evidence:** explicit WpdMtpBackend(bool selfTestWpdCacheBackend) noexcept : _selfTestWpdCacheBackend(selfTestWpdCacheBackend) {}
// if (_selfTestWpdCacheBackend) { constexpr std::string_view payload = "RedSalamander deterministic MTP fixture\r\n"; ... }  — in release code paths

**Suggested fix:** Wrap the fixture members and branches in #ifdef _DEBUG alongside the factory, or inject the enumerate/open/find primitives so the cache logic can be tested against FakeMtpBackend without a fixture mode in the production class.


## Test Quality (8)

### 34. [HIGH][test-quality] CompareDirectories RunCase count assertion is off by one on master AND in the working tree; the correcting fix is stranded in unmerged snapshot commit 31a0b6b83

**Location:** `Tools/Tests/TestInventory.Tests.ps1:35`  |  **Found by:** merge-integrity (+tests-tools)

Commit da6b438a0 ('chore: checkpoint test stabilization work') added 5 new SelfTest::RunCase registrations to RedSalamander/SelfTest/CompareDirectories (230 -> 235 actual) but bumped the inventory assertion and the coverage doc only to 234 (+4). Result: on master, Get-RSTestInventory returns 235 while TestInventory.Tests.ps1 expects 234, so BOTH assertions fail deterministically (line 35 exact-count check and line 76 doc-vs-source check: 'CompareDirectories static RunCase count drifted' / 'Coverage spec CompareDirectories count drifted from source'). The fix (234 -> 235) exists ONLY in off-master commit 31a0b6b83 ('Codex worktree snapshot: archive-cleanup', 2026-07-05, child of master head da6b438a0, reachable from no branch) — the exact 'fix silently lost off master' failure mode this audit targets. Worse, the uncommitted Operation Granite working tree re-derived its bump from the stale committed value: it adds 3 new RunCase registrations (+1 Cases.Mtp.cpp, +2 Cases.SearchAndIndex.cpp) and asserts 234+3=237, but the true count is 235+3=238 (verified by executing Get-RSTestInventory against the working tree: returns 238; per-file grep: CoreDiffs 48 + Mtp 51 + RuntimeAndRemote 37 + SearchAndIndex 102 = 238). So committing the Granite work as-is perpetuates the failing test.

**Evidence:** Working tree Tools/Tests/TestInventory.Tests.ps1:35: "Assert-RSEqual -Actual $inventory.SelfTests.CompareDirectories.RunCaseRegistrations -Expected 237" while Get-RSTestInventory -RepoRoot Z:\src\RedSalamander returns 238. Per-commit actual/asserted counts: 53e7e6641 230/230, e814abbb6 230/230, da6b438a0 (master) 235/234, working tree 238/237. git show 31a0b6b83 contains only the 234->235 correction in Testing_TestCoverage.md and TestInventory.Tests.ps1 and is on no branch. da6b438a0 added exactly 5 '+SelfTest::RunCase(options,' lines to the CompareDirectories case files.

**Suggested fix:** Set the expectation to 238 in Tools/Tests/TestInventory.Tests.ps1:35 and update Specs/Testing/Testing_TestCoverage.md (lines 32 and 2008) from 237 to 238; then delete/ignore the superseded snapshot 31a0b6b83. Consider making the doc the single source (the line-76 doc-vs-source assertion already exists) instead of hardcoding the count twice.


### 35. [MEDIUM][test-quality] OK/Cancel 'live UIA InvokePattern' assertions are now satisfiable by a debug-hook keyboard path, silently removing the UIA coverage the test claims to verify.

**Location:** `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp:1494`  |  **Found by:** selftests-a

invokeButtonOrActivateTarget returns true if either the UIA invoke works OR DebugFocusCompareDirectoriesOptionsTargetForWindow + sendSpaceToTargetHost works. The second alternative uses only in-process debug hooks and a synthetic Space key — no UIA involvement at all. If InvokePattern support for the DX footer buttons regresses completely (the exact regression the original Require text 'did not expose live UIA InvokePattern interaction' was written to catch), the test now passes via the backdoor with no trace that the fallback was taken. Unlike the named-button relaxation nearby (which still requires UIA-derived namedButtonControlCount), this fallback fully bypasses the accessibility surface.

**Evidence:** CompareOptions.cpp:1494-1495: return InvokeVisibleDescendantByNameWithMessagePump(compare, UIA_ButtonControlTypeId, buttonText, label) || (DebugFocusCompareDirectoriesOptionsTargetForWindow(compare, target) && sendSpaceToTargetHost(target, label)); used at 1498 for Cancel and at the OK require (formerly UIA-only asserts).

**Suggested fix:** At minimum Trace loudly when the keyboard fallback path is used (so regressions surface in run logs), or keep a separate non-fatal UIA-only Require alongside the combined one.


### 36. [MEDIUM][test-quality] New worker_shutdown test scrapes product source text and even asserts member declaration order — the same fragility this very diff had to patch in three other files

**Location:** `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.RuntimeAndRemote.cpp:1775`  |  **Found by:** selftests-b

The behavioral first half of worker_shutdown_joins_before_state_teardown (create session, observe worker activity, session.reset()) genuinely exercises destructor ordering. But the second half reads CompareDirectoriesEngine.cpp/.h off disk and asserts exact code strings: `joinWorkers(_scanWorkers);`, `_contentCompareQueueNotFullCv.notify_all();` ordering, and that the literal text `_scanWorkers` appears before `_contentCompareCache` in the header. Any behavior-preserving refactor (rename, reorder, clang-format re-wrap) fails the test. This diff itself demonstrates the failure mode: it adds three duplicated `compactWhitespace` lambdas (Cases.SearchAndIndex.cpp:2606 and 2946, Commands.SelfTest.PluginConfig.cpp:4874) purely to un-break older source-scrape assertions after formatting changed `x = y;` spacing. Adding a fourth scrape site (with an even stricter declaration-order assertion) compounds the maintenance trap, and these assertions also require the source checkout to exist at the __FILE__-derived path at test run time.

**Evidence:** state.Require(contains(destructorBody, "joinWorkers(_scanWorkers);") && contains(destructorBody, "joinWorkers(_contentCompareWorkers);"), ...);
state.Require(containsBefore(engineHeader, "_scanWorkers", "_contentCompareCache"),
              L"CompareDirectoriesSession scan workers are declared before content cache; explicit destructor joins are required.");

**Suggested fix:** Keep the behavioral half; drop the source-text assertions (or at minimum hoist one shared whitespace-insensitive Contains helper into SelfTest common instead of the three per-file compactWhitespace copies this diff adds).


### 37. [LOW][test-quality] CollectVisibleCompareOptionsEditDiagnostics' timeout is illusory: request_stop() targets a worker lambda that ignores its stop_token, and the outer loop still blocks until the worker finishes.

**Location:** `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp:184`  |  **Found by:** selftests-a

The worker lambda is declared '[sharedState, hwnd](std::stop_token) noexcept' (line 133) — the token is unnamed and never checked — so worker.request_stop() at line 184 has no effect. Unlike RunUiaActionWithMessagePump, which detaches and returns on timeout, the outer 'while (!sharedState->done)' keeps pumping until the worker completes regardless of the deadline; timedOut only discards the collected result. If the UIA enumeration stalls, this diagnostics helper stalls the whole case with it, and the trace message ('timed out') misleadingly implies the wait was bounded.

**Evidence:** CompareOptions.cpp:133: std::jthread worker([sharedState, hwnd](std::stop_token) noexcept  — token unnamed;  CompareOptions.cpp:183-184: timedOut = true; worker.request_stop();  followed by the loop continuing until sharedState->done.

**Suggested fix:** Since the enumeration is a single non-cancellable pass, drop the request_stop/timedOut machinery and just pump until done (returning whatever was collected), or mirror RunUiaActionWithMessagePump's detach-on-timeout.


### 38. [LOW][test-quality] New failure-path test checks shortcut non-existence with plain std::filesystem::exists while sibling checks in the same change were migrated to \\?\-prefixed GetFileAttributesW, risking a vacuous pass.

**Location:** `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp:3679`  |  **Found by:** selftests-a

This same diff replaced std::filesystem::exists-based shortcut checks with QueryShellShortcutPathForShellCommandTest (GetFileAttributesW on BuildShellPathForShellCommandTest's \\?\ path) and even routed IPersistFile::Load through the shell path (ShellCommands.cpp:2320-2321), evidently because plain paths misbehave on the deep GUID-suffixed suite temp tree. The new TestPaneClipboardPasteShortcutFailureAfterNavigateShowsAlert asserts the negative — '!exists(staleLink)' — with the plain API: if exists() errors out (e.g., path-length), it returns false and the assertion passes vacuously, so a forced-failure worker that DID leave a stray .lnk behind would go undetected.

**Evidence:** ShellCommands.cpp:3679: state.Require(! std::filesystem::exists(staleLink, ec), L"Forced Paste Shortcut failure should not leave a shortcut in the stale destination.");  vs. the positive checks in the same file: bool alphaLinkExists = false; const HRESULT alphaExists = QueryShellShortcutPathForShellCommandTest(alphaLink, alphaLinkExists); state.Require(SUCCEEDED(alphaExists) && alphaLinkExists, ...).

**Suggested fix:** Use QueryShellShortcutPathForShellCommandTest(staleLink, exists) and require SUCCEEDED(hr) && !exists, matching the positive-existence checks.


### 39. [LOW][test-quality] Secure-clear contract assertion matches the bare substring "true)" anywhere in the accessor block, so the guarded flag can regress undetected

**Location:** `Tests/DxUiTests/DxUiTests.TextField.cpp:1539`  |  **Found by:** tests-tools

The assertion intends to verify TextField::GetOrCreateSingleLineLayout passes secureClear=true to GetOrCreateSingleLineTextLayout (DxUi.TextInput.cpp:3149), but it only requires that "true)" appears somewhere in the block. If the trailing 'true' argument is removed (regressing secure wipe of cached password text on layout failure) while any other call in the block ends with 'true)' — e.g. a future 'Invalidate(host, true)' — the guard still passes. Every neighboring assertion in this test matches the full call text (e.g. "ClearSingleLineTextLayoutCache(_singleLineLayoutCache, true)"), so this one is anomalously weak for a security-relevant contract.

**Evidence:** Require(accessorBlock.find("true)") != std::string_view::npos,
        "TextField single-line accessor asks the shared helper to securely clear cached text on layout failure");

**Suggested fix:** Match the meaningful text, e.g. accessorBlock.find("readingDirection, true)") or the full GetOrCreateSingleLineTextLayout(...) call tail, mirroring the ClearSingleLineTextLayoutCache assertions.


### 40. [LOW][test-quality] ViewerVLC focus check is now self-fulfilling: the test sets focus on the video window itself before asserting 'routes keyboard focus into the video surface'

**Location:** `Tests/ViewerPETests/ViewerPETests.cpp:4648`  |  **Found by:** tests-tools

The old code asserted the product routed initial focus into a visible child after SetFocus(viewerWindow) ('ViewerVLC routes initial focus into a visible child surface'). The new code, on timeout, calls SetFocus(videoWindow) directly and re-polls, then Checks videoFocusReady with a message claiming the viewer routes focus. A regression where ViewerVLC stops forwarding initial focus to its video child now passes silently (the test forces the focus itself), and the SetFocus(viewerWindow) result check was dropped too. Only the Tab-to-HUD transfer remains a genuine product assertion. This is presumably deliberate flake mitigation, but the check message no longer matches what is verified and the initial-focus-routing contract is unguarded.

**Evidence:** static_cast<void>(SetFocus(viewerWindow));
bool videoFocusReady = PumpUntil([&]() noexcept { return GetFocus() == videoWindow; }, 2000ms);
if (! videoFocusReady)
{
    static_cast<void>(SetFocus(videoWindow));
    videoFocusReady = PumpUntil([&]() noexcept { return GetFocus() == videoWindow; }, 2000ms);
}
Check(videoFocusReady, L"ViewerVLC routes keyboard focus into the video surface", success);

**Suggested fix:** Either keep a separate (possibly diagnostic-only, non-latching) check for product-driven initial focus routing and reword the latching check to 'video surface can take keyboard focus', or record when the fallback path was taken so routing regressions remain visible.


### 41. [LOW][test-quality] Resource-file discovery switched from exclude-list to a hardcoded root allowlist, so .rc files in any future top-level directory silently escape localization contract checks

**Location:** `Tools/Tests/ResourceLocalizationContracts.Tests.ps1:104`  |  **Found by:** tests-tools

Previously Get-RSResourceFiles scanned every repo root child except .build/packages/.claude, so a new top-level project with .rc resources was automatically covered. Now only 7 named roots are scanned (Common, Plugins, PoC, RedConfigure, RedSalamander, RedSalamanderMonitor, Tests). Adding a new root project (e.g. a new tool/app dir) with .rc files that violate the localization contracts will pass this suite with no signal. Today the allowlist matches all real .rc locations, so the regression is latent, not current.

**Evidence:** $resourceRoots = @(
    'Common',
    'Plugins',
    'PoC',
    'RedConfigure',
    'RedSalamander',
    'RedSalamanderMonitor',
    'Tests'
)

**Suggested fix:** Add a guard test that fails if any *.rc exists under a repo-root child not in the allowlist (still excluding .build/packages/.claude), so new roots must be consciously added.


## Simplifications (20)

### 42. [MEDIUM][simplification] The 4-line 'GetHost -> RefreshWindowHostAccessibilitySnapshot' block is copy-pasted 13 times across DxUi, and Grid's refresh was made public so app code must manually call it after every selection-model mutation

**Location:** `Common/DxUi/DxUi.cpp:143`  |  **Found by:** arch-lens

This diff open-codes the identical host-fetch-and-refresh block in Control::SetVisible/SetEnabled/SetFocusable/SetAccessibleName/SetAccessibleHelpText (DxUi.cpp:143, 162, 179, 439, 457), Label::SetText/SetMnemonicTarget, Button::SetText, Toggle::SetChecked (DxUi.Controls.cpp:1756, 1803, 1848, 2404), and adds private per-class duplicates Grid::RefreshAccessibilitySnapshot (DxUi.Grid.cpp:1239) and Tree::RefreshAccessibilitySnapshot (DxUi.Tree.cpp:360) — 13 copies of one pattern. Separately, Grid::RefreshAccessibilitySnapshot was exposed publicly (DxUi.h:2970) and app code now has to remember to call it after every GetSelectionModel().SetSingle()/Clear() — 11 new manual call sites in Preferences.Plugins.cpp (846, 858, 867, 886, 898, 907, 1523, 1542, 1571) and Preferences.Keyboard.cpp (826, 1291). Any future selection-mutation site that forgets the call silently ships a stale accessibility snapshot. A single protected Control::RefreshAccessibilitySnapshot() removes all 13 copies, and having GridSelectionModel notify its owning Grid on mutation removes the entire manual app-side call surface.

**Evidence:** if (WindowHost* const host = GetHost())
{
    RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
}
// verbatim at DxUi.cpp:143/162/179/439/457, DxUi.Controls.cpp:1756/1803/1848/2404, DxUi.Grid.cpp:1239, DxUi.Tree.cpp:360, DxUi.ComboBox.cpp:2293, DxUi.TextInput.cpp:949

**Suggested fix:** Add protected Control::RefreshAccessibilitySnapshot() and route selection-model mutations through the Grid (observer or wrapper methods) instead of a public refresh that callers must remember.


### 43. [MEDIUM][simplification] RenameItem adds a third near-verbatim copy of the ~40-line overwrite orchestration lambda already present in CopyItem and MoveItem

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:3320`  |  **Found by:** arch-lens

The uncommitted RenameItem change (Core.cpp:3320-3380) duplicates the destination-probe/policy/temp-swap sequence from CopyItem (Core.cpp:3103-3143) and MoveItem (Core.cpp:3199-3239): GetAttributes probe, missing->S_OK mapping, directory->ERROR_ACCESS_DENIED, createdObjectPuidUnsupported->ERROR_NOT_SUPPORTED with the same perf counter, CommitDeviceSourceOverwriteWithTempSwap on existing destination, plain backend op otherwise, plus the identical tempPuidMissing/tempPuidPresent probe-recording epilogue. The only deltas are the same-path ERROR_ALREADY_EXISTS check and which backend fallback is invoked. Three hand-synced copies of overwrite/journal policy is exactly the kind of drift risk this data-safety code cannot afford — a fix to one copy (e.g. a new policy check) will silently miss the other two.

**Evidence:** RenameItem: 'commandHr = backend.GetAttributes(commandDest, destinationAttributes); ... CommitDeviceSourceOverwriteWithTempSwap(backend, commandSource, commandDest, true, verifyLevel, journalContext, tempPuidMissing, tempPuidPresent);' — same block modulo one flag at 3103-3143 (CopyItem) and 3199-3239 (MoveItem)

**Suggested fix:** Extract a helper taking (backend, source, dest, moveSemantics, fallback op) that performs the probe/policy/temp-swap orchestration, and call it from all three methods.


### 44. [MEDIUM][simplification] setAndWaitForEditValue duplicates a full 3-second wait and its nested waits exceed the outer retry deadline, so the SetValue retry loop can never actually retry and a failing set burns ~9.5s.

**Location:** `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp:1394`  |  **Found by:** selftests-a

waitForEditValue (CompareOptions.cpp:1365-1383) polls with its own 3s deadline plus a final read. Inside setAndWaitForEditValue's do-while (outer deadline also 3s, line 1387), a failed attempt executes 'setNamedEditValue(...) && waitForEditValue(...)' (up to 3s) and then immediately 'if (waitForEditValue(expectedValue))' again (another full 3s) for the same condition — by which time the outer deadline has expired, so the loop exits after a single iteration and SetValue is never re-issued (the retry machinery is dead). Then line 1401 runs a third 3s waitForEditValue and lines 1406-1417 a 500ms confirmation loop: ~9.5s+ per failing set, and the function is called four times in this test (~40s worst case). In a suite with a documented multi-day convergence problem this both inflates wall time and hides that the intended set-retry never happens.

**Evidence:** CompareOptions.cpp:1390-1397: if (setNamedEditValue(expectedValue, label) && waitForEditValue(expectedValue)) { return true; } if (waitForEditValue(expectedValue)) { return true; }  — where waitForEditValue at 1367 has 'deadline = now + SelfTest::Scale(3000ms)' equal to the outer deadline at 1387.

**Suggested fix:** Give waitForEditValue a short per-attempt timeout parameter (e.g., 250-500ms) inside the retry loop, delete the duplicated second waitForEditValue call, and keep one final full-length wait after the loop.


### 45. [LOW][simplification] Dead machinery in DeactivateNativeTextInputTsf: shouldReleasePreviousFocusDocumentAfterPop guard is a no-op and Disconnect+Detach are duplicate calls

**Location:** `Common/DxUi/DxUi.NativeTextInput.cpp:829`  |  **Found by:** dxui

(1) `shouldReleasePreviousFocusDocumentAfterPop` is computed, then used to conditionally reset _nativeTextInputTsfPreviousFocusDocumentMgr at line 853-856 — but the same member is reset unconditionally 20 lines later (line 876) with no reads in between, and resetting an empty com_ptr is already a no-op, so the bool and its if-block are dead code. (2) DetachNativeTextInputTextStore just forwards to Disconnect() (DxUi.TextStoreACP.cpp:177-180: `void DetachHost() noexcept { Disconnect(); }`), so calling DisconnectNativeTextInputTextStore (line 887) followed by DetachNativeTextInputTextStore (line 895) on the same store performs the identical operation twice; one call suffices.

**Evidence:** const bool shouldReleasePreviousFocusDocumentAfterPop = static_cast<bool>(_nativeTextInputTsfPreviousFocusDocumentMgr); ... if (shouldReleasePreviousFocusDocumentAfterPop) { _nativeTextInputTsfPreviousFocusDocumentMgr.reset(); } ... _nativeTextInputTsfPreviousFocusDocumentMgr.reset();  — and DisconnectNativeTextInputTextStore(textStoreToDisconnect.get()); ... DetachNativeTextInputTextStore(textStoreToDisconnect.get());

**Suggested fix:** Delete the bool and its conditional reset; keep a single Disconnect call (or collapse Detach/Disconnect into one function since they are now identical).


### 46. [LOW][simplification] Debug trace scaffolding (18+ TraceNativeTextInputTsfStep call sites) left in production TSF activate/deactivate paths, tripling their length and obscuring the logic

**Location:** `Common/DxUi/DxUi.NativeTextInput.cpp:35`  |  **Found by:** dxui

ActivateNativeTextInputTsf/DeactivateNativeTextInputTsf grew from ~30 lines of logic each to ~100+ lines dominated by step-by-step trace calls (activate.enter, activate.push.before/after, deactivate.set-focus-null.before/after, ...), each re-reading the REDSALAMANDER_DXUI_TEXTINPUT_TRACE_FILE env var and re-opening the file per line when enabled. This was clearly diagnostic instrumentation for the flake investigation ('checkpoint test stabilization work') that shipped to master. The env-var gate makes it functionally harmless, but the ordering-sensitive TSF teardown logic — the part that actually needs review — is now buried in tracing noise, and the file-per-line write pattern is wasteful if anyone ever enables it.

**Evidence:** TraceNativeTextInputTsfStep(L"activate.push.before", ...); hr = documentMgr->Push(context.get()); TraceNativeTextInputTsfStep(L"activate.push.after", ...);  (repeated ~18 times across the two functions)

**Suggested fix:** Either remove the scaffolding now that the investigation checkpoint passed, or collapse to a single scoped tracer object (cache the env-var lookup once, one call site per operation) so the control flow reads as TSF logic again.


### 47. [LOW][simplification] Unreachable final return in AuthorizeClientRootRebuildAccess; lastMissingError's initializer is never used

**Location:** `Common/SearchServiceBroker.cpp:2518`  |  **Found by:** search

The while loop can only be exited via one of the three return statements inside it: normalizedRoot is checked non-empty at entry (line 2480-2483), and authorizationPath is only reassigned from parentPath after an explicit parentPath.empty() check returns (lines 2509-2513), so the loop condition '! authorizationPath.empty()' never becomes false and the trailing 'return HRESULT_FROM_WIN32(lastMissingError);' at line 2518 is dead. Likewise the initializer 'lastMissingError = ERROR_PATH_NOT_FOUND' (line 2493) is always overwritten (line 2508) before any read. The dead tail suggests a reachable exit path that does not exist, which will mislead future maintainers of this security-sensitive walk-up.

**Evidence:** std::wstring parentPath = ParentPathForAuthorization(authorizationPath);
if (parentPath.empty() || OrdinalString::EqualsNoCase(parentPath, authorizationPath))
{
    return HRESULT_FROM_WIN32(lastMissingError);
}

authorizationPath = std::move(parentPath);
}

return HRESULT_FROM_WIN32(lastMissingError);   // line 2518 — unreachable

**Suggested fix:** Use 'for (;;)' and delete the trailing return (or keep the loop and mark the tail std::unreachable()); drop the unused ERROR_PATH_NOT_FOUND initializer.


### 48. [LOW][simplification] deviceIoMutex is now dead serialization machinery: it is only ever locked by the single worker thread of the queue that exclusively owns it

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:2001`  |  **Found by:** mtp-core

After the command-queue refactor, the only remaining lock of *deviceIoMutex is inside MtpBackendCommandQueue::WorkerMain (line 2001). Each queue has exactly one worker thread, and AbandonBackendSessionLocked always allocates a fresh mutex (line 2236) before any new queue could be constructed, so no two threads can ever contend on the same mutex instance — the queue itself already provides the per-instance transport serialization the spec requires. The shared_ptr<std::mutex> member, the State field, and the per-command lock_guard can all be removed without any behavior change, which also removes the misleading impression that something else still synchronizes on it.

**Evidence:** std::lock_guard deviceLock(*state->deviceIoMutex);   // WorkerMain, sole locker
...
void FileSystemMtp::AbandonBackendSessionLocked(...)
{
    _deviceIoMutex = std::make_shared<std::mutex>();   // fresh mutex per session generation; old one stays exclusive to the quarantined worker

**Suggested fix:** Delete _deviceIoMutex, the State::deviceIoMutex field, and the WorkerMain lock; keep a comment that the single worker thread is the serialization point.


### 49. [LOW][simplification] WriteOverwriteJournalReplayAttempt is now a one-line pass-through wrapper with a single caller

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:1094`  |  **Found by:** mtp-core (+arch-lens)

The refactor extracted WriteOverwriteJournalEntry and left WriteOverwriteJournalReplayAttempt as a trivial forwarding function used only once (line 1515), while the new retained-replay path calls WriteOverwriteJournalEntry directly (line 1379). The wrapper adds a second name for the same operation and no behavior.

**Evidence:** [[nodiscard]] HRESULT WriteOverwriteJournalReplayAttempt(const OverwriteJournalContext& context,
                                                         const OverwriteJournalEntry& entry,
                                                         uint32_t replayAttemptCount) noexcept
{
    return WriteOverwriteJournalEntry(context, entry, replayAttemptCount);
}

**Suggested fix:** Call WriteOverwriteJournalEntry directly at line 1515 and delete the wrapper.


### 50. [LOW][simplification] WpdMemoryBackendFileReader and FakeMtpStreamingReader are byte-for-byte duplicate implementations of the same in-memory IMtpBackendFileReader

**Location:** `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:1027`  |  **Found by:** arch-lens

WpdMemoryBackendFileReader (FileSystemMtp.Device.cpp:1027-1122) and FakeMtpStreamingReader (FileSystemMtp.FakeBackend.cpp:253-360) implement identical GetSize/Seek/Read logic over a std::vector<std::byte>, down to the same non-obvious negative-offset handling ('static_cast<uint64_t>(-(offset + 1)) + 1u') and overflow checks; the fake adds only a stats hook and an optional sleep. Both also emit the same 'mtp.transfer.read_bytes' counter. Any future fix to the seek/overflow arithmetic must be applied twice.

**Evidence:** const uint64_t delta = static_cast<uint64_t>(-(offset + 1)) + 1u;
if (base < delta) { return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK); }
// identical in FileSystemMtp.Device.cpp:1058-1064 and FileSystemMtp.FakeBackend.cpp:284-292

**Suggested fix:** Move one MemoryBackendFileReader (with optional per-read hook for stats/delay) next to IMtpBackendFileReader in FileSystemMtp.Internal.h and reuse it from both backends.


### 51. [LOW][simplification] Dead unreachable `case WM_NCDESTROY:` left behind after the early-intercept refactor

**Location:** `RedSalamander/ConnectionManagerWindow.cpp:4730`  |  **Found by:** folderwindow-nav

WindowProc now intercepts WM_NCDESTROY at the top of the function (lines 4692-4696) and returns before the message switch, so the `case WM_NCDESTROY: OnNcDestroy(); return 0;` in the switch can never execute. Leaving both paths suggests OnNcDestroy could run twice and obscures the intent of the new early intercept.

**Evidence:** Line 4692-4696: `if (msg == WM_NCDESTROY) { OnNcDestroy(); return 0; }` ... line 4730: `case WM_NCDESTROY: OnNcDestroy(); return 0;`

**Suggested fix:** Delete the unreachable `case WM_NCDESTROY:` from the switch.


### 52. [LOW][simplification] Hide-issues-pane focus-restore block is duplicated verbatim in two files

**Location:** `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp:898`  |  **Found by:** folderview-fileops

The same 12-line sequence (capture GetFocus, test focused==pane||IsChild, ShowWindow(SW_HIDE), then RequestRestoreFolderViewFocus(GetFolderViewHwnd(GetActivePane())) with TryRestoreActivePaneFolderViewFocus() fallback) was added twice by this change: FileOperationsIssuesPaneState::OnClose (IssuesPane.cpp:889-909) and FolderWindow::FileOperationState::ToggleIssuesPane (State.Diagnostics.Part.cpp:221-237). Both operate on the same FolderWindow; future focus-policy tweaks (e.g., choosing the pane the pane-window was opened from instead of the active pane) must be made in two places or the two close paths diverge.

**Evidence:** IssuesPane.cpp:898-909: `if (restoreFolderFocus && folderWindow && hostLifetime.lock()) { const HWND folderView = folderWindow->GetFolderViewHwnd(folderWindow->GetActivePane()); if (folderView && IsWindow(folderView) != FALSE) { folderWindow->RequestRestoreFolderViewFocus(folderView); } else { static_cast<void>(folderWindow->TryRestoreActivePaneFolderViewFocus()); } }` — identical (modulo `_owner.` vs `folderWindow->`) to State.Diagnostics.Part.cpp:226-237.

**Suggested fix:** Extract a FolderWindow helper, e.g. `void RestoreActivePaneFolderViewFocusAfterChildHide() noexcept`, and call it from both sites.


### 53. [LOW][simplification] Issues-pane hide + folder-view focus-restore block duplicated between ToggleIssuesPane and the pane's OnClose, with a latent SaveViewState divergence

**Location:** `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp:221`  |  **Found by:** arch-lens

ToggleIssuesPane (State.Diagnostics.Part.cpp:221-237) and FileOperationsIssuesPaneState::OnClose (FolderWindow.FileOperations.IssuesPane.cpp:889-909) now contain the same new logic: capture GetFocus(), test focused==pane||IsChild, SaveIssuesPanePlacement, ShowWindow(SW_HIDE), then GetFolderViewHwnd(GetActivePane()) -> RequestRestoreFolderViewFocus else TryRestoreActivePaneFolderViewFocus. Besides the duplication, the two hide paths already diverge: OnClose also calls SaveViewState() while the ToggleIssuesPane copy does not, so hiding via toggle silently loses grid view state that hiding via close preserves — and future edits to the focus-restore policy must be made twice.

**Evidence:** State.Diagnostics.Part.cpp:221-237 and IssuesPane.cpp:889-909 contain the identical focusedBeforeHide/IsChild/RequestRestoreFolderViewFocus block; only OnClose additionally runs 'SaveViewState();'

**Suggested fix:** Have ToggleIssuesPane hide by sending the pane WM_CLOSE (reusing OnClose), or extract one helper that saves state, hides, and restores folder-view focus.


### 54. [LOW][simplification] The 'post prompt close debug command' helper is implemented three times verbatim in three translation units

**Location:** `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp:3863`  |  **Found by:** arch-lens

PostDxUiPromptCloseDebugCommand (FolderWindow.FileSystem.Commands.Part.cpp:3863), PostFolderViewPromptCloseDebugCommand (FolderWindow.FileSystem.Navigation.Part.cpp:853), and PostFileOperationsSpeedLimitPromptCloseDebugCommand (FolderWindow.FileOperations.Popup.cpp:8118) are the same function (validate hwnd/message, PostMessageW with the same rationale comment about not tearing down TSF prompts inside a cross-thread SendMessage). Ten Debug*Prompt call sites were rewritten to use one of the three copies. One shared helper in a common ENABLE_TESTS header removes two copies and keeps the SendMessage->PostMessage teardown policy in one place.

**Evidence:** if (! hwnd || message == 0u || IsWindow(hwnd) == FALSE) { return false; }
// Confirm/Cancel destroys modal DxUi HWNDs. Post the close action so selftest workers do not tear down TSF/native text input inside a cross-thread synchronous SendMessage stack.
return PostMessageW(hwnd, message, command, 0) != FALSE;  // duplicated at Navigation.Part.cpp:853 and Popup.cpp:8118

**Suggested fix:** Hoist a single PostDxUiPromptCloseDebugCommand into a shared internal header used by all three files.


### 55. [LOW][simplification] PostFolderViewPromptCloseDebugCommand duplicates PostDxUiPromptCloseDebugCommand but drops the message==0 guard

**Location:** `RedSalamander/FolderWindow.FileSystem.Navigation.Part.cpp:853`  |  **Found by:** folderwindow-nav

The diff adds two file-local helpers with identical purpose: PostDxUiPromptCloseDebugCommand (FolderWindow.FileSystem.Commands.Part.cpp:3863) validates `message == 0u` before posting, while PostFolderViewPromptCloseDebugCommand does not. Two of its call sites pass RegisterWindowMessageW results (GetFolderViewCreateDirectoryPromptDebugMessage / GetFolderViewEditNewPromptDebugMessage, FolderWindow.FileSystem.cpp:881-891) which return 0 on registration failure; PostMessageW(hwnd, 0 /*WM_NULL*/, ...) then succeeds, so the helper reports true and the selftest waits out its full WaitForWindowClosed timeout and fails with a misleading 'prompt did not close' instead of 'failed to confirm'. Merging the two helpers into one shared ENABLE_TESTS utility with the message!=0 guard removes the duplication and the inconsistency.

**Evidence:** Navigation.Part.cpp:853-864: `if (! hwnd || IsWindow(hwnd) == FALSE) { return false; } ... return PostMessageW(hwnd, message, command, 0) != FALSE;` vs Commands.Part.cpp:3863-3874: `if (! hwnd || message == 0u || IsWindow(hwnd) == FALSE) { return false; }`

**Suggested fix:** Share a single helper (with the `message == 0u` check) between the two translation units.


### 56. [LOW][simplification] Third verbatim copy of the dropdown focus-restore block

**Location:** `RedSalamander/NavigationView.Menus.cpp:2359`  |  **Found by:** folderwindow-nav

The new block appended to ShowDiskInfoDropdown (check _requestFolderViewFocusCallback, GetAncestor(GA_ROOT), conditional SetActiveWindow, invoke callback) is an exact copy of the blocks already at lines 2116-2124 (history dropdown) and 2627-2634. Three copies of the same 8-line activation/focus dance is a maintenance trap: the next tweak to the activation condition (e.g. skipping the restore when the executed action opened external UI such as the drive Properties sheet or cleanmgr from FileSystem::ExecuteDriveMenuCommand) has to be applied in three places.

**Evidence:** Lines 2359-2367: `if (_requestFolderViewFocusCallback) { const HWND root = GetAncestor(_hWnd.get(), GA_ROOT); if (root && GetActiveWindow() != root) { SetActiveWindow(root); } _requestFolderViewFocusCallback(); }` — identical to lines 2116-2124 and 2627-2634.

**Suggested fix:** Extract a private NavigationView::RestoreFolderViewFocusAfterDropdown() helper and call it from all three dropdown handlers.


### 57. [LOW][simplification] Compare-pane RestoreDeferredFocusAfterLayout() was added to the Plugins deferred-action branch instead of the CompareDirectoriesIgnoreToggleChanged branch, making it dead code in the wrong case

**Location:** `RedSalamander/Preferences.Dialog.cpp:5221`  |  **Found by:** compare-prefs

The new call to hostState._compareDirectoriesPane.RestoreDeferredFocusAfterLayout() sits inside the PluginsSearchChanged/PluginsConfigure/PluginsTest/PluginsTestAll case of HandleDeferredPaneAction. The compare pane's deferred focus target is already consumed by LayoutPreferencesPageHost's compare-directories branch (Preferences.Dialog.cpp:3788, added in the same diff), which is what runs for the actual CompareDirectoriesIgnoreToggleChanged case at line 5236-5245 (that case has no direct call). While the Plugins page is active, _compareDirectoriesPane._deferredFocusAfterLayout is None so the call is a no-op; in the narrow stale case (a compare toggle set the member but the posted deferred action was dropped, e.g. PostMessagePayload failure) a later Plugins-page action would consume/act on compare-pane focus state while a different page is showing. It reads as a copy-paste into the wrong case block; either way it is redundant with line 3788.

**Evidence:** case PreferencesDeferredActionKind::PluginsSearchChanged:
case PreferencesDeferredActionKind::PluginsConfigure:
case PreferencesDeferredActionKind::PluginsTest:
case PreferencesDeferredActionKind::PluginsTestAll:
{
    const bool handled = hostState._pluginsPane.HandleDeferredAction(host, state, payload.kind);
    if (handled && host)
    {
        LayoutPreferencesPageHost(host, state);
        InvalidateRect(host, nullptr, FALSE);
        hostState._compareDirectoriesPane.RestoreDeferredFocusAfterLayout();
    }
    return handled;
} (Preferences.Dialog.cpp:5211-5224); compare branch already covered: hostState._compareDirectoriesPane.RestoreDeferredFocusAfterLayout(); (Preferences.Dialog.cpp:3788)

**Suggested fix:** Delete the call at line 5221 (the CompareDirectoriesIgnoreToggleChanged path is already handled inside LayoutPreferencesPageHost); additionally clear _deferredFocusAfterLayout in CompareDirectoriesPane::DetachDxHosts to prevent stale targets surviving a pane teardown.


### 58. [LOW][simplification] This diff adds four more copies of extractJsonUInt (8 total) with divergent overflow handling, plus substring counter assertions that match wrong values

**Location:** `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp:1905`  |  **Found by:** selftests-b

The new cases each define their own extractJsonUInt lambda (lines 1905, 2170, 2253, 5914 are new; 8 copies now exist in this one file) and two more copies of the plugin-private stableDeviceHash (now 12 copies). The copy at line 1905 (mtp_reader_streams_on_read_not_open) silently lacks the multiply-overflow guard the other copies have. Meanwhile other new assertions bypass the parser entirely and use raw substring matching: `props.find(R"json("copyItemCalls":1)json")` (lines 6192, 6811) also matches copyItemCalls values 10-19, and the same pattern is used for maxConcurrentBackendCalls. A shared numeric-JSON-field helper in the selftest common header would make every counter assertion exact and delete ~150 duplicated lines.

**Evidence:** line 1905 copy: value = (value * 10u) + static_cast<uint64_t>(json[position] - '0');  // no overflow guard
line 2170 copy: if (value > ((std::numeric_limits<uint64_t>::max)() - digit) / 10u) { return std::nullopt; }
line 6192: state.Require(props.find(R"json("copyItemCalls":1)json") != std::string_view::npos, ...)  // also matches ":10"..":19"

**Suggested fix:** Hoist one extractJsonUInt (with overflow guard) and stableDeviceHash into the CompareDirectories selftest common scope, and assert `extractJsonUInt(props, "copyItemCalls") == 1u` instead of substring find.


### 59. [LOW][simplification] New self-test case duplicates ~130 lines of helper lambdas already copied five times in the same file

**Location:** `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp:5783`  |  **Found by:** merge-integrity

The uncommitted new case mtp_overwrite_journal_clears_completed_swap_without_temp (working tree, starts line 5783) inlines yet another verbatim copy of the helper lambdas narrowAscii, stableDeviceHash, getLocalAppDataPath, ensureDirectoryExists, and writeUtf8File. The file already contains five identical copies (at lines ~3840, ~4082, ~4340, ~4656, ~4931 per grep), and this hunk adds a sixth. Payoff: hoisting these to file-scope static helpers removes ~650 duplicated lines across the six journal-replay cases and eliminates the risk of the copies drifting (e.g., stableDeviceHash must stay bit-identical to the plugin's OverwriteJournalDeviceIdentity hashing for the journal-path computation to be valid — six independent copies of that FNV variant is a correctness trap).

**Evidence:** grep -n 'stableDeviceHash\|getLocalAppDataPath\|ensureDirectoryExists\|writeUtf8File' CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp shows definitions repeated at lines 3840/4082/4340/4656/4931 and again inside the new case added by the working-tree diff (hunk '@@ -5753,6 +5778,292 @@').

**Suggested fix:** Extract the five lambdas into anonymous-namespace free functions at the top of the file (or a shared self-test helper header) and have all six journal cases call them; no behavior change.


### 60. [LOW][simplification] Unreachable 'continue' branches after Require, which calls std::exit(1) on failure

**Location:** `Tests/DxUiTests/DxUiTests.NativeTextInput.cpp:198`  |  **Found by:** tests-tools

In TestFolderViewDxUiPromptsDeactivateNativeTextInputBeforeDestroyWindow, three 'if (... == npos) { continue; }' blocks (lines 198-201, 239-242, 270-273) immediately follow requireMessage(...) calls asserting the same condition is not npos. Require (DxUiTestHelpers.h:29) prints and std::exit(1)s on failure, so the continue branches can never execute. They read as if the loop tolerates missing blocks when it actually hard-fails, which is misleading dead code.

**Evidence:** requireMessage(closeHelper != std::string_view::npos, std::string(promptClasses[index]) + " centralizes modal HWND close");
if (closeHelper == std::string_view::npos)
{
    continue;
}

**Suggested fix:** Delete the three unreachable if/continue blocks (or, if graceful continuation is desired, make requireMessage non-fatal for these checks — pick one behavior).


### 61. [LOW][simplification] Invoke-RSShowPerfRunsCommand duplicates ~40 lines of Invoke-RSShowPerfRuns (pwsh discovery, output redirection, Start-Process wrapper), and the suite gates on the live repo budget file

**Location:** `Tools/Tests/ShowPerfRuns.Tests.ps1:146`  |  **Found by:** tests-tools

The two helpers differ only in the pwsh argument list (-File script args vs -Command text); everything else (Get-Command pwsh check, process-output dir, guid stdout/stderr files, Start-Process flags, result object) is copy-pasted. Additionally, because Show-PerfRuns hardcodes the budget path from $PSScriptRoot, the tests can only exercise budget minimums against the real Specs\Testing\FolderViewPerfBudgets.json5 — 'uses FolderView budget minimumSamples...' hardcodes the expectation that folder.frame.input_to_paint_us has minimumSamples 40, so routine budget retuning breaks this test with no product change.

**Evidence:** function Invoke-RSShowPerfRunsCommand { ... $pwshCommand = Get-Command pwsh -ErrorAction SilentlyContinue ... $process = Start-Process -FilePath $pwshCommand.Source -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $CommandText) ... } — near-identical body to Invoke-RSShowPerfRuns at line 91

**Suggested fix:** Extract a single Invoke-RSPwsh helper taking the argument array; consider a -BudgetPath override parameter on Show-PerfRuns.ps1 so tests can pin a synthetic budget file instead of the live one.



## Manually verified findings (verifier fleet hit session limit; confirmed by hand)

### 62. [MEDIUM][bug] Validation-fallback rename desyncs `_selectedConnectionName`: Connect saves the new profile name but reports/connects the stale one

**Location:** `RedSalamander/ConnectionManagerWindow.cpp:4162`  |  **Found by:** folderwindow-nav, verified manually

The new `ResolveEditedModelIndexForValidation()` (added in this diff) lets `TryValidateAndNormalizeConnectionProfiles()` locate the edited profile by `_selectedConnectionName` when the grid has **no** selection, and then rename `_connections[index].name` from the editor text (line 3082). Back in `OnConnectClicked()` (line 4160), `_selectedConnectionName` is refreshed from the model **only** when `GetSelectedModelIndex()` has a value — exactly the case where the fallback did *not* fire. When the fallback fired, `_selectedConnectionName` keeps the old name: `SaveConnectionsSettings()` persists the renamed profile, then `_modalResult->connectionName` (line 4175) and `NotifyOwnerToConnectSelectedProfile()` (payload built from `_selectedConnectionName`, line 3321) both carry a name that no longer exists in settings. Secondary inconsistency: `OnCloseClicked()` clears `_selectedConnectionName` *before* validating (lines 4117-4118), so the same editor rename is silently dropped on the Close path.

**Suggested fix:** after `TryValidateAndNormalizeConnectionProfiles()`, refresh `_selectedConnectionName` via `ResolveEditedModelIndexForValidation()` (or have validation return the resolved index) instead of `GetSelectedModelIndex()` alone.

### 63. [MEDIUM][bug] Per-root service rejections poison the instance-wide 5s service-unavailable cooldown and mislabel the warning

**Location:** `Plugins/FileSystem/FileSystem.Search.cpp:2805`  |  **Found by:** search, verified manually

The GR-13 fallback (intentional: degrade to local INDEX/SCAN when the service refuses a root) also executes the *global* cooldown arm: any fallback-candidate HRESULT — including per-request root rejections (`ERROR_BAD_PATHNAME`, `E_INVALIDARG`) from a perfectly healthy service — sets `_searchServiceUnavailablePipeName`/`_searchServiceUnavailableUntilTick` (lines 2803-2807). For the next 5 s **all** searches, including ones on valid indexable roots, skip the service (lines 2660-2661 → 2705-2707) and are labeled `FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE` (lines 2711-2714) even though the service is up. The per-request fallback and warning for *that* query are by design; the cross-query cooldown poisoning and its "service unavailable" label are collateral.

**Suggested fix:** arm the cooldown only for transport/connect-class failures (pipe missing, timeout), not for per-request validation rejections; or use a distinct warning flag for root-rejected-by-service.

### 64. [LOW][test-quality] Raw-provider fallback drops the "visible" contract: no `UIA_IsOffscreenPropertyId` check

**Location:** `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp:3980`  |  **Found by:** selftests-a, verified manually

The primary UIA lookup requires `UIA_IsOffscreenPropertyId == FALSE` (condition built at line 3899); the raw fallback `RawProviderMatchesVisibleDescendant` (lines 3980-4004) matches on control type + name only, so it can match an offscreen element the primary path would reject. All four call sites are on the worker-thread paths already confirmed dead by finding #9 — this becomes load-bearing the moment that finding is fixed, so fix both together (add an `IsOffscreen` read via the existing `TryReadRawProviderLongProperty`-style helper).

## Refuted findings (appendix)

Three findings were killed in adversarial verification; recorded here so they are not re-reported by future reviews:

1. **`DxUi.Tree.cpp:871` — "Tree::OnKeyDown steals Win32 focus"**: unreachable in production. Queued `WM_KEYDOWN` is assigned to the focus window at `GetMessage` time, so organic input guarantees `GetFocus()==hwnd`; the only production `DxUi::Tree` is the Preferences category tree, and the mid-handler rebuild path already restores focus to the same window. Remains a (low) hygiene note: test-stabilization side effect living in production input handling, inconsistent with Grid.
2. **`FileSystem.Search.cpp:1961` — "E_INVALIDARG as blanket fallback trigger swallows future client bugs"**: intentional, documented, RED/GREEN-tested GR-13 behavior; the debug selftest explicitly asserts `IsServiceFallbackCandidate(E_INVALIDARG)`. Current behavior correct; hypothetical-future-bug argument, and the proposed fixes have concrete downsides (wire-contract skew would reintroduce the GR-13 hard-fail).
3. **`CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp:10808` — "rebuild test deletes the live service's own storage root"**: impossible with `--store-backend=snapshot`: every store write is synchronous within the triggering RPC (temp write → flush → close → rename), there are no async writers for the snapshot backend, and the assertion tests the in-memory index, not the on-disk snapshot.

Two further "unverified" entries were duplicates of confirmed findings and were folded into them: `FileSystemMtp.Device.cpp:1552` (COM lifetime across `CoInitializeEx` cycles → findings #8/#13) and `FileSystemMtp.Core.cpp:1134` (overwrite-journal absent-cache cross-instance race → finding #12).
