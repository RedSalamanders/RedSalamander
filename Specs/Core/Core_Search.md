# Search Specification

## Overview

RedSalamander search is a host-owned feature that spans the command system, the Find dialog, optional plugin-native search, a generic host scan fallback, the built-in local index core, and the Windows search service backend.

This document is the authoritative architecture and behavior spec for that system.

Related normative documents:
- `Specs/Plugins/Plugins_VirtualFileSystem.md` for the `IFileSystemSearch` ABI and capability JSON.
- `Specs/UI/UI_CommandMenuKeyboard.md` for `cmd/pane/find` command routing and shortcut defaults.
- `Specs/UI/UI_FolderWindow.md` for pane ownership, focus routing, and modeless window integration.
- `Specs/Core/Core_SettingsStore.md` for persisted Find dialog state.
- `Specs/Testing/Testing_SelfTests.md` for self-test result and skip semantics.
- `Specs/Testing/Testing_PerformanceValidation.md` for mandatory performance-validation requirements.

Historical implementation detail and rollout evidence remain in:
- `Specs/Plans/Done/Core_Search_PluginArchitecturePlan.md`
- `Specs/Plans/Done/RFC_Core_InstantFileSearchUsnMftIndex.md`

## Goals

- Provide one Find experience for all filesystem plugins.
- Prefer fast local indexed search when the active pane is browsing a supported local root.
- Preserve correctness when native or indexed search is unavailable by falling back to a generic scan path.
- Keep the UI responsive by running all searches off the UI thread.
- Match Windows local-filesystem name semantics for case-insensitive search: case-insensitive, accent-sensitive, and normalization-sensitive unless the underlying filesystem itself behaves differently.
- Make degradation visible instead of silently changing behavior.

## Performance Validation Contract

Search is a performance-sensitive subsystem. Any new search feature, backend change, optimization, or behavior change that can affect responsiveness, queueing, batching, rebuild frequency, fallback routing, or visible latency MUST:

- define the protected scenario up front,
- add or reuse measurable instrumentation,
- add deterministic selftest coverage or a deterministic harness,
- archive validation runs under `Specs/TestRuns/`,
- report baseline vs candidate evidence when claiming improvement.

This is a normative requirement, not a later cleanup step.

Search work SHOULD prefer the existing `find.ui.*` metrics and the relevant `--commands-selftest` / `--compare-selftest` paths when those already cover the scenario; otherwise the change MUST add the missing coverage with the feature.

## Non-Goals

- No separate search-plugin type in v1.
- No synthetic `search:` virtual filesystem in v1.
- No binary or hex content search in the Find dialog in v1.
- No service-side file-content indexing or content transport in v1.

## Command and Window Behavior

### `cmd/pane/find`

- `cmd/pane/find` opens a new independent modeless `Find Files and Directories` window.
- The command targets the focused pane when focus is inside a pane; otherwise it targets the active pane.
- The default shortcuts are `Alt+F7` and `Ctrl+F`.
- Multiple Find windows may be open at the same time; each keeps its own search context and result set.

### Default scope

The initial search scope is derived from the target pane:
- active filesystem plugin,
- active instance context,
- current pane path as `rootPath`.

Persisted `lastRoot` and recent root history are suggestions/history only; they MUST NOT override the target pane path when the Find command opens or reactivates the window. Reinvoking `cmd/pane/find` on an existing idle Find window refreshes `Look in` from the current target pane path.

### Find window responsibilities

The Find window is an independent top-level, theme-aware tool window. It must:
- run on the existing DirectX-rendered themed host control stack,
- stream results incrementally,
- show backend, progress, warning, completion, and backend database or synchronization state,
- preserve keyboard/focus behavior consistent with the rest of the app,
- persist placement and recent inputs through the shared settings store.

## Query Semantics

- At least one of name matching or content matching must be enabled.
- If both name and content matching are enabled, a result must satisfy both.
- Name matching supports wildcard, literal, and regex modes.
- Content matching supports text literal and text regex modes.
- v1 content search is text-oriented only.
- `FILESYSTEM_SEARCH_WANT_SNIPPETS` is a hint, not a guarantee.
- `FILESYSTEM_SEARCH_PREFER_INDEX` is a hint, not a guarantee.
- `FILESYSTEM_SEARCH_FORCE_SCAN` is a host override for implementations that support multiple local backends. It takes precedence over `FILESYSTEM_SEARCH_PREFER_INDEX` and backend `auto` policy and must route the request to live scan when scan is supported.
- Name and content regex patterns MUST be validated through the shared `SearchTextHelpers::ValidateRegexPatternSafety(...)` policy before traversal or content reads begin, on both native and host-fallback paths. Syntax errors and safety rejects (for example nested unbounded repetition, overlapping top-level alternatives such as `(a|a)*`/`(a|ab)*x`, or excessive nesting/length) MUST fail with `E_INVALIDARG`, emit no matches, and report exactly one forced `FILESYSTEM_SEARCH_PHASE_COMPLETED` progress payload with `FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED` and `statusHint == E_INVALIDARG`. Provably bounded quantifiers such as `{n}` and `{n,m}` MUST NOT be rejected merely because they are quantified.
- Regex evaluation is cooperatively cancellable before and after each candidate evaluation; no code path may promise interruption inside a single `std::regex_search` call.
- Decoded regex-content inputs larger than `SearchTextHelpers::kMaxRegexContentCharacters` MUST NOT become silent clean misses. Implementations may either match a documented capped prefix and report `FILESYSTEM_SEARCH_WARNING_OVERFLOW`, or skip the match and report `FILESYSTEM_SEARCH_WARNING_OVERFLOW`.
- A rejected regex search MUST leave no active worker/search state behind. For `IFileSystemSearch::Search(...)`, callbacks remain per-call and MUST NOT be invoked after the method returns.

## Result Model

### Set operations

The Find window supports four result-set operations:
- `Find`: replace the current result set.
- `Append`: union the new result set into the current result set.
- `Intersect`: keep only items present in both sets.
- `Subtract`: remove newly found items from the current result set.

Find/Append may stream visible result mutations as matching rows arrive. Intersect/Subtract MUST be transactional: each search has an epoch, result/progress/completion payloads from stale epochs MUST be ignored, and matched keys for the active set operation MUST be staged separately from the visible result set. The staged Intersect/Subtract operation commits exactly once only after the current epoch completes with a clean `S_OK` and the user did not cancel or abandon the operation. Failed, cancelled, or stale Intersect/Subtract operations MUST preserve the existing visible result set and clear only the staged keys. A successful zero-match Intersect is a valid empty intersection and MUST clear the result set.

The visible Find action is a split button: its primary action runs `Find`, and its menu exposes `Find`, `Intersect`, `Subtract`, and `Append`. Set operations other than `Find` MUST remain disabled while the current result set is empty, and MUST become available once at least one result is present. When the split menu opens under a stationary pointer, the item under the pointer MUST receive the normal hover highlight without requiring extra mouse movement. While that menu is open, the owning split button MUST keep its pressed/highlighted visual state until the menu closes.

Find dialog editable combos, including `Look in`, use the shared single-line editing contract: repeated `Ctrl+Backspace` deletes successive previous word/path segments, and translated control characters generated by editing shortcuts MUST be ignored rather than inserted into the field. Find request construction and persisted combo history MUST strip single-line control characters so stale invisible text cannot change visible wildcard patterns such as `*`.

`Go to folder` on a result MUST navigate the focused pane to the result parent and focus the result entry. When the focused pane is already showing that parent folder and the item is already present, this action MUST focus the item directly without forcing a full pane refresh.

Find result `Open` and `Go to folder` actions initiated from the dialog controls MUST target the pane returned by `FolderWindow::GetFocusedPane()` at action execution time. A physical mouse click on `Go to folder` for a result whose parent differs from the focused pane MUST complete pane navigation and enumeration without requiring any later pointer movement, including when the Find search root differs from the currently focused pane folder. Pressing plain `Space` while the Find results grid is focused MUST leave selection unchanged; `Enter` continues to invoke `Open`.

Find result rows MUST render real shell bitmap icons for local files and folders through `IconCache`; Fluent text glyphs are fallback only when a shell icon bitmap cannot be resolved. The icon index is the system image-list index for the file extension, directory class, special folder, or per-file path as appropriate. One-line result rows in compact mode MUST visually match `FolderView` brief rows: 24 DIP effective row height, 16 DIP list icon slot, and the DxUI `ListItem` text role. Runtime compact-density changes MUST relayout the results grid immediately through the Find grid metrics path.

`Go to folder` must return to the message loop promptly while any required pane navigation/enumeration continues through the pane's normal asynchronous navigation path.

Find-driven pane navigation MUST NOT synchronously query NavigationView disk information on the UI thread. The active pane may queue `IDriveInfo::GetDriveInfo` through NavigationView's async worker; that worker-side latency is reported separately as `navigation.drive_info.query_us` and must not delay the `Go to folder` click handler returning to the message loop.

### Canonical result identity

Result identity is frozen as:
- `pluginId + normalized instanceContext + normalized fullPath`

This identity is used by the Find window for de-duplication and set algebra across native, indexed, service, and fallback backends.

The normalization step for result identity MUST use the shared invariant search-folding helpers (`OrdinalString::FoldCaseInvariant` and helpers built on it) for the plugin ID, instance context, and full path. It MUST NOT use user/thread-locale-sensitive folding such as `towlower` or direct `CharLowerBuffW` call sites.

### Result actions

- Double-clicking a file opens it through the existing host open/view flow.
- Double-clicking a directory navigates the pane to that directory.
- The secondary parent action navigates to the parent directory and focuses the matched item when possible.
- Result-grid copy/move-to-other-pane commands must show the resolved destination folder in the Find status line after the shared File Operations task is accepted.
- Result-grid viewer/editor commands must use the Find window as the action owner and must not steal focus by routing through a main-pane folder-view command post.
- Recursive results MUST display the containing subfolder relative to the search root in the Path column, including when a backend reports only a leaf relative path but provides a full path under the requested root.
- One-line result rows MUST shrink when compact mode is active and align to the FolderView brief-row visual contract; snippet rows may remain taller to preserve preview readability.

## Execution Model

### SearchSessionController

The host runs searches through a dedicated session controller that owns:
- worker-thread lifetime,
- cancellation,
- backend selection,
- result batching,
- progress aggregation,
- completion delivery.

`IFileSystemSearch::Search(...)` remains synchronous, so the host must always invoke it on a worker thread.

### UI-thread contract

- The UI thread must never block on directory traversal, content reads, index warm-up, or service calls.
- Worker results are marshaled back to the Find window through posted messages.
- Closing a Find window while a search worker is active MUST request cancellation, hide/disable the modeless window, and return to the message pump without joining the worker. The window may be destroyed only after the current worker has reached its quiet point and the matching completion payload has settled on the UI thread.
- Starting a new Find search MUST NOT rely on `std::jthread` move-assignment to clean up a previous worker; any join must be explicit and limited to a worker already proven idle.
- Wildcard/name-only scans MUST surface incremental result batches while the search is still active, including large `*` searches over a single directory. The Find window MUST not rely on low-priority timers for the first visible batch, because result floods can starve timers and make the window appear blank until completion.
- Cancellation must be checked at traversal boundaries and during chunked content reads.
- Built-in local parallel name-only and content scans MUST also check cancellation while draining completed worker result batches. A large directory that produces many ready matches MUST NOT defer `FileSystemSearchShouldCancel` until the whole batch has been emitted.
- Indexed backends (`LocalSearchIndexCore::Repository::Query`/`Enumerate`) and broker query streams MUST honor a real `CancelCheckFn` during candidate enumeration and between streamed service batches, returning `ERROR_CANCELLED` with bounded partial results instead of completing a full result set after cancellation becomes visible.
- `IFileSystemSearch` callbacks for a single `Search(...)` call MUST be serialized even when parallel scan workers finish concurrently. The coordinator may abandon queued completed batches after a terminal cancellation/failure once all workers have reached the quiet point.
- Parallel recursive scan MUST avoid starvation: broad sibling directories and deep descendant chains both have to make progress unless the search is cancelled or a terminal error occurs.
- Recursive scan traversal MUST queue/visit subdirectories before applying the result name mask to that directory entry. A folder that does not match the current mask can still contain matching descendants and MUST NOT be pruned solely because its own name fails the mask.
- Built-in local scan metadata construction MUST avoid normalizing/materializing full and relative child paths for file entries until the path is needed for recursive directory traversal, content evaluation, or an emitted match. Name-only scans MUST match against display names first, and non-matching file entries SHOULD keep zero path materializations in deterministic coverage. `FileSystem.Search.EntryPathMaterialized` records the materialized-path-pair count versus scanned entries for regression detection.

## Backend Architecture

### Backend precedence

For the built-in local `file` plugin, backend selection order is:
- `service`
- `local-index`
- `scan`

The configuration override values are:
- `auto`
- `service`
- `local-index`
- `scan`

`auto` uses the precedence above.

When the Find dialog runs against an explicit built-in local context and `Prefer indexed backend` is unchecked, it MUST set `FILESYSTEM_SEARCH_FORCE_SCAN`; unchecked means "scan now", not "let auto choose service or local-index".

### Generic host scan fallback

The host scan fallback is the correctness baseline for all plugins. It uses:
- `IFileSystem::ReadDirectoryInfo` for traversal,
- `IFileSystemIO::CreateFileReader` for content scanning when available.

Fallback behavior:
- If a plugin does not expose `IFileSystemSearch`, the host may still provide name search through scan fallback.
- Content fallback additionally requires `IFileSystemIO`.
- If no readable stream API exists, content search must degrade cleanly and surface `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT`.
- Built-in local scan/search paths MUST use the same canonical local path contract as File Operations: ordinary drive/UNC Win32 paths and supported extended forms (`\\?\C:\...`, `\\?\UNC\server\share\...`) are normalized with slash folding and Win32 `.` / `..` collapse before traversal identity, containment, and watch/index decisions; generic device namespaces such as `\\?\GLOBALROOT\...`, `\\.\...`, and `\??\...` are not accepted as ordinary local search roots by prefix rewriting. Display casing remains separate from folded comparison keys.
- When `FILESYSTEM_SEARCH_FOLLOW_SYMLINKS` is set, host fallback and native local traversal MUST remain bounded through directory symlink/junction/reparse loops. Local Win32 paths MUST deduplicate followed directory targets by physical identity `(volume serial number, file index)` in addition to folded logical path. If the physical identity probe fails for a followed local directory, skip that directory, roll back the logical visit key, and surface `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED`. Generic plugin fallback paths that cannot expose physical identity MUST enforce hard recursion-depth and queued-directory caps and surface `FILESYSTEM_SEARCH_WARNING_OVERFLOW` when a cap skips traversal. Native local follow-symlink traversal MUST also cap retained queued-directory identity state and surface `FILESYSTEM_SEARCH_WARNING_OVERFLOW` when the cap prevents further traversal.
- Directory traversal MUST iterate each returned `IFilesInformation` buffer once via `GetBuffer()`/`GetBufferSize()` and `NextEntryOffset`. The host fallback traversal path MUST NOT enumerate entries with `GetCount()` + `Get(index)` because plugin convenience accessors may locate each index from the buffer head.
- Malformed `FileInfo` records from `ReadDirectoryInfo` MUST NOT cause out-of-bounds reads or abort the whole search when safe recovery is possible. If an entry has invalid name bounds or an odd byte length but a valid `NextEntryOffset`, skip that entry, continue with siblings, and surface `FILESYSTEM_SEARCH_WARNING_OVERFLOW`. If the traversal link itself is corrupt or points outside the buffer, stop only that directory's buffer walk, keep prior results, and surface the same warning.
- Native local search paths that cannot materialize an otherwise matched entry path (for example, an empty display name from a malformed directory entry) MUST skip only that entry, continue with later siblings/chunk entries, and surface `FILESYSTEM_SEARCH_WARNING_OVERFLOW`. Parallel workers MUST NOT abandon the remainder of a chunk for a materialization failure. Parallel name-only search MUST flush already accumulated matches even if cancellation is observed before the directory result is otherwise clean, but flushing matches MUST NOT overwrite or hide a fatal directory enumeration status; fatal status must still propagate to the terminal search result so truncated results are not reported as a clean success.

### Built-in local indexed backend

The shared local index core is used by both the built-in `file` plugin and the Windows service backend.

Supported indexed roots:
- NTFS
- ReFS

Unsupported roots must fall back to scan:
- FAT / exFAT
- UNC
- WSL-backed paths
- other roots where required journal or enumeration support is unavailable

The local index core is responsible for:
- path normalization,
- snapshot persistence for non-service callers and sqlite persistence for service-authoritative roots,
- query planning,
- journal replay,
- tracked-root traversal rebuilds when direct journal access is unavailable.

Local index rebuild/enumeration MUST ignore RedSalamander staged sibling names (`.rs_tmp_*`, `.~rs-write-*`, and `.rs_copy_tmp_*`) so aborted, crashed, or in-flight local and cross-filesystem staging temps do not become searchable user documents. The filter MUST recognize only generated RedSalamander temp-name shapes, such as bridge `.rs_tmp_` plus generated hex suffixes, copy `.rs_copy_tmp_<pid>_<tid>_<counter>` hex suffixes, and writer `.~rs-write-<guid-or-hex-id>.tmp` suffixes. It MUST NOT suppress arbitrary user files merely because a marker substring appears in the middle of an otherwise normal name.

Snapshot persistence MUST write a sibling temporary file, validate every write byte count, flush the temporary handle, and replace the final snapshot with a write-through atomic rename only after the full snapshot has been written. Directory creation, size inspection, final and temporary opens, cleanup, atomic replacement, reload, deletion, and test corruption MUST retain extended-length path support; the generated sibling name can cross `MAX_PATH` even when the configured snapshot root and final filename do not. Failed, partial, or crashed snapshot saves MUST leave the previous final snapshot intact. Snapshot persistence `noexcept` helpers that translate non-allocation C++ exceptions to `HRESULT` MUST document the boundary, log once with `Debug::Error(...)`, and keep `std::bad_alloc` fatal.

### Allocation and protocol safety

Search persistence, journal, and service decoders MUST treat counts and byte lengths from snapshots, USN records, service frames, plugin buffers, or any other external/corruptible source as untrusted. They MUST validate those values against the remaining byte span or an explicit configured ceiling before `reserve`, `resize`, string assignment, or per-entry loops can allocate from them.

`SearchServiceBroker::Query` callers that do not supply `CandidateBatchCallbackFn` MUST be protected by an explicit client-side buffered-candidate ceiling before decoded batches are appended to `outCandidates`. The ceiling MUST consider both candidate count and estimated retained candidate storage, and `request.maxResults` MUST lower that ceiling when it is set. If a streamed response would exceed the ceiling, the broker returns `ERROR_BUFFER_OVERFLOW` and clears partial buffered candidates. Callers that need large streams should supply `CandidateBatchCallbackFn` and consume batches without retaining the full result set.

Corrupt local snapshots return `ERROR_INVALID_DATA` and rebuild. Malformed service frames return a failing `HRESULT_FROM_WIN32(RPC_S_PROTOCOL_ERROR)` so normal `FAILED(hr)` fallback paths degrade to a slower backend instead of treating the protocol error as success. Malformed USN records are skipped or rejected before reading `FileName` data. True `std::bad_alloc` remains fatal and MUST NOT be swallowed by broad recovery catches.

### Journal replay delta contract

Journal replay is a transaction over tracked file IDs, not copied path snapshots. Replay MUST:
- treat `FILE_DELETE` and `RENAME_OLD_NAME` as subtree removals and record removed IDs in the persisted delta,
- upsert a `FILE_CREATE`, `RENAME_NEW_NAME`, basic-info, hard-link, or reparse-point change only when the record is still reachable from the tracked root, or when it is the tracked root itself,
- remove a previously tracked entry when its new parent is outside the tracked tree,
- track moved-in/new directories by `NodeId`, rebuild derived path state, then hydrate each still-tracked directory from its fresh indexed path so pre-existing children become searchable without a full root rebuild,
- for SQLite-authoritative roots, apply journal delta deletes before upserts and include all hydrated descendants in `upsertIds`.

If the USN journal becomes invalid during replay after an initial currentness check (`ERROR_JOURNAL_ENTRY_DELETED`, `ERROR_JOURNAL_NOT_ACTIVE`, or `ERROR_JOURNAL_DELETE_IN_PROGRESS` from `FSCTL_READ_USN_JOURNAL`), the index MUST treat that as an invalid journal range, rebuild the tracked root, and report `rebuiltJournalRangeInvalid` instead of surfacing a hard query failure.

Regression coverage must exercise snapshot and SQLite-authoritative stores for populated directory move-in, extracted/unzipped tree move-in, move-out deletion, and in-tree rename. The expected proof is exact indexed-result equality plus persisted SQLite row equality, with no NTFS-enumeration fallback for the replay.

### Windows service backend

The Windows search service is the preferred local backend when available.

Build-specific default identities:
- Debug: service name `RedSalamanderSearchService.Debug`, pipe `\\.\pipe\RedSalamander.SearchService.Debug.v3`
- Release: service name `RedSalamanderSearchService`, pipe `\\.\pipe\RedSalamander.SearchService.v3`

Service rules:
- communicate through a versioned named-pipe protocol,
- use build-specific default storage roots:
  - Debug: `%ProgramData%\\RedSalamander\\SearchIndex.Debug`
  - Release: `%ProgramData%\\RedSalamander\\SearchIndex`
- persist the default SQLite store as `index-v2.sqlite3` under the active storage root unless `--sqlite-path` overrides it,
- keep those default SQLite roots behaviorally isolated on disk: starting a Debug service under the default SQLite configuration must not create or mutate the Release root, and vice versa,
- when a build-specific default SQLite store already exists at that default root, reuse it in place without migration, rename, or sibling-root creation,
- keep the SQLite schema at version `2` with:
  - `meta(key, value)` for store metadata such as `schema_version`, `store_kind`, `store_generation`, `last_checkpoint_utc`, and `last_compaction_utc`,
  - `volumes(volume_id, root_path, fs_kind, journal_id, next_usn, state, entry_count, last_seed_utc, last_replay_utc, last_error_hr)` for per-root indexing state,
  - `entries(volume_id, file_id_low, file_id_high, parent_id_low, parent_id_high, full_path, full_path_folded, name, name_folded, extension_folded, attributes, is_dir, size_bytes, write_time_100ns, creation_time_100ns, last_access_time_100ns, change_time_100ns, allocation_size)` for indexed filesystem metadata,
- persist indexed filesystem metadata, root journal state, and maintenance timestamps, but not file contents, query history, snippets, or transient in-memory synchronization/query-progress state,
- treat the SQLite store as the authoritative persisted current state plus journal cursor for service-managed roots,
- accept status and query requests before startup warm-up, rebuild, repair, or SQLite mirroring completes,
- derive query cutover readiness from every persisted volume state: only `Ready` volumes are query-ready. `ImportedLegacySnapshot` remains the separate legacy-migration counter, while `CurrentnessUnproven` blocks query cutover without being reported as legacy-import work,
- choose the fastest no-wait query path per request:
  - direct SQLite only when the configured store is valid and current for the requested root,
  - live filesystem scan fallback otherwise,
- never block the request thread on `EnsureReady`, rebuild, repair, or SQLite mirroring,
- keep service request-thread routing strictly two-path in sqlite-authoritative mode:
  - direct SQLite only when currentness is provably up to date for the requested root,
  - live filesystem scan fallback otherwise while background warm-up or repair continues,
- keep direct SQLite read paths (`InspectVolume`, `LoadVolume`, and `EnumerateVolume`) strictly read-only. They must open the existing store with a read-only connection, validate that schema v2 is already present, and must not bootstrap, migrate, write metadata, or take a writer transaction while answering a query.
- persist a monotonically increasing SQLite `store_generation` in `meta`. Store bootstrap and v2 migration must initialize the key when it is missing, and every committed `ReplaceVolume`, `ApplyJournalDelta`, and existing-volume `DeleteVolume` transaction must increment it. Cached `PersistentStoreInfo` used by query paths must carry the inspected generation and validate it against the configured SQLite store before opening/querying the cached store. Query validation may skip opening SQLite when the cached and current database/WAL `(mtime,size,exists)` stamps are unchanged; creation or retirement of an empty read-only WAL sidecar with an unchanged database stamp is also steady state. Any database change or non-empty/newly changed WAL must probe the generation, and one `Enumerate` operation must reuse its validated store result rather than probing twice. A mismatch or unreadable generation must refresh store info before the query chooses direct SQLite or live fallback. External store rotation must therefore be detected within the next query, before the query relies on the `search.backend.sqlite.retry_query_ms` path.
- status construction must not overlay a coherent on-disk inspection with cached repository counts from a different `store_generation`; a mismatch is treated conservatively as possible external rotation and clears generation-specific retained store/sync/fallback/root state plus the prior query execution mode. If the same configured store is merely uninspectable, status may retain how the last request executed, but it must re-derive store/sync/fallback/root state from the failed inspection and independent warm-up snapshot. A retained Ready/Watching snapshot must not override a newly inspected cutover-blocked store, and an old syncing/cutover-blocked snapshot must not mask a newly inspected Ready generation.
- direct SQLite volume lookup is case-insensitive under the same invariant Windows-style folding used by search matching. Exact stored-root text remains the display/root persistence casing; differently-cased callers must find the stored volume without creating a duplicate root row.
- implement SQLite name-prefix prefiltering as a folded half-open `name_folded` range (`>= prefix` and `< prefixUpper`) over `idx_entries_name_folded`; do not use `LIKE` for prefix prefilters. User `%` and `_` characters in wildcard prefix patterns are literal filename characters, not SQLite wildcards.
- automatic SQLite maintenance MUST detect stores whose persisted `PRAGMA auto_vacuum` mode is not `INCREMENTAL`, queue idle maintenance for them even without unrelated WAL or freelist pressure, and run the required `PRAGMA auto_vacuum=INCREMENTAL` + full `VACUUM` rewrite outside bootstrap before relying on incremental vacuum for later space reclamation,
- automatic SQLite maintenance MUST avoid redundant checkpoint passes. Checkpoint-only, full-VACUUM rewrite, and incremental-vacuum maintenance should dirty maintenance metadata first and then perform one final truncate checkpoint that reflects the actual post-maintenance WAL shape,
- full SQLite `VACUUM` maintenance (manual compaction and automatic auto-vacuum-mode rewrite) MUST run `PRAGMA quick_check` first and abort with a failure if the store is not clean. Background checkpoint-only maintenance that hits `SQLITE_BUSY`/`SQLITE_LOCKED` may defer with `S_FALSE`; explicit/manual and post-VACUUM checkpoint failures remain hard maintenance failures.
- `last_checkpoint_utc` MUST advance only after the final checkpoint succeeds. A background checkpoint deferred with `S_FALSE` must leave the prior timestamp unchanged. SQLite enumeration observability counts only candidates accepted by the callback in `emittedRows`; the internal callback-skip sentinel does not increment it.
- when NTFS journal/MFT access is unavailable because the service is not elevated, reuse the same traversal seed strategy used for ReFS, persist that traversal seed into SQLite as currentness-unproven, and keep query cutover degraded until a later run can prove currentness from the live journal,
- keep file content out of the service protocol in v1,
- report missing, invalid, stale, and cutover-blocked SQLite state by degrading to live filesystem search with `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX`,
- if a direct SQLite query hits missing, invalid, stale, or cutover-blocked SQLite state, refresh persistent-store info and retry as a belt-and-suspenders path even when early hits have already been emitted; emitted candidate path keys must be deduped across the retry and any live fallback instead of surfacing duplicates or silently degrading,
- use invariant Windows-style case folding through the shared `OrdinalString` helpers for indexed and fallback path/name matching, literal content matching, result identity, and folding-only in-pane incremental search, so `A.txt` and `a.txt` match the same entry while composed and decomposed Unicode names remain distinct,
- model each indexed file ID with one canonical reachable path. NTFS hardlinks and other multiple-path aliases are therefore not guaranteed exhaustive in indexed results; when seed or replay observes distinct paths for one file ID, `QueryStats::hardlinkAliasCoverageIncomplete` MUST be set and brokered query stats MUST preserve that flag. Directory junction/mountpoint hydration follows a coverage-only alias contract: it must traverse an alias subtree, with physical `NodeId` cycle protection, when that is required to discover descendants not present during the canonical-path seed. Alias-only descendants must be indexed once through the canonical node model; emitting both canonical and alias result paths requires a future explicit alias-path representation.
- fall back safely to `local-index` or `scan` when the service is absent, mismatched, disconnected, or unavailable,
- service root-validation errors that the local scanner can still handle, including `ERROR_BAD_PATHNAME`, `ERROR_PATH_NOT_FOUND`, and `E_INVALIDARG`, MUST be treated as client fallback candidates rather than hard Find failures,
- arm the client-wide service-unavailable cooldown only for transport/connect failures. A healthy service's rejection of one root must fall back with `FILESYSTEM_SEARCH_WARNING_SERVICE_ROOT_REJECTED`, must not claim `FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE`, and must not prevent an immediate valid-root query from using the service,
- keep client named-pipe operations cancelable and bounded. Status, query, rebuild, compact, and foreground-shutdown requests MUST use overlapped client I/O with per-frame and per-operation deadlines; query I/O MUST poll the request cancellation callback while waiting, including missing-pipe connection retries, and call `CancelIoEx` on cancellation or timeout. A stalled or absent service must produce a prompt cancellation or fallback-eligible failure instead of pinning a search worker indefinitely.

Service trust-boundary and pipe protocol rules:
- the pipe is local-machine only (`PIPE_REJECT_REMOTE_CLIENTS`) and its ACL must grant full access only to LocalSystem and Builtin Administrators, with non-admin interactive clients limited to read/write pipe access,
- the service must treat every client-supplied root as untrusted input. Query and rebuild roots MUST be rejected when empty, relative, UNC, generic device namespace (`\\.\`, `\??\`), `\\?\UNC`, `\\?\GLOBALROOT`, or otherwise not a canonical drive-rooted path accepted by the local index contract,
- accepted client roots MUST be normalized before use. The service may strip a `\\?\` prefix only for drive-rooted paths, must collapse `.`/`..` through the Win32 full-path resolver, and must preserve drive-root spelling without widening the request to a device namespace,
- before a service-identity query acts on a client root, the server MUST impersonate the named-pipe client and prove the client can open that exact root with the minimum required directory/file read attributes and traversal/list/read access. Failed impersonation or failed access checks reject only that client request and must not terminate the service,
- rebuild/invalidate requests use the same root normalization and trust-boundary rejection rules, but they MUST be allowed to purge a root that no longer exists. When the requested normalized root is missing with `ERROR_FILE_NOT_FOUND` or `ERROR_PATH_NOT_FOUND`, the server MUST impersonate the client, ACL-check the deepest existing ancestor, keep rejecting genuine access-denied or malformed roots, and invalidate the original requested root rather than widening the purge to the ancestor,
- per-candidate authorization failures after a query has been accepted MUST NOT abort the whole query. The service must skip the affected candidate, keep previously validated candidates, and report `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED` through both streamed batch headers and final query stats/runtime warnings from one per-query accumulator. Candidate-directory authorization caches may retain `false` across queries only for durable negative results such as `ERROR_ACCESS_DENIED`, `ERROR_FILE_NOT_FOUND`, and `ERROR_PATH_NOT_FOUND`. A transient `CreateFileW` failure such as `ERROR_SHARING_VIOLATION`, `ERROR_BAD_NETPATH`, or resource exhaustion may be cached and logged once per parent within the current query/batch to avoid repeated failing opens, but must not poison a later query,
- service request handling MUST use one consistent overlapped pipe I/O model. Server reads/writes must be bounded by per-frame deadlines and the service stop event, must cancel outstanding I/O with `CancelIoEx` on timeout/shutdown, and must treat zero-byte transfers, short frames, oversized frames, bad message types, and protocol-version mismatches as protocol failures,
- explicit protocol shutdown is enabled only for foreground-mode servers whose owning harness/process selected its unique pipe. SCM-hosted service instances MUST reject that request; their lifetime remains owned by the Service Control Manager,
- a malformed, unauthorized, disconnected, or slow client MUST be disconnected after the failed request while the foreground/service process continues accepting later clients. Only non-recoverable server failures may stop the service process.

Service executable CLI:
- `RedSalamanderSearchService.exe --help` prints the supported command-line options.
- `--run-foreground` runs the service directly in the current terminal instead of through the Service Control Manager.
  It auto-attaches to the parent terminal when needed, prints a startup banner with PID/build/mode details, renders a full-screen `FTXUI` dashboard in a real VT-capable console, and emits readable lifecycle log lines when stdout/stderr are redirected.
  The interactive dashboard must:
  - use the terminal alternate screen and stay full-page,
  - wake on service state changes instead of repainting on a blind timer when the service is idle,
  - keep a lightweight pulse animation only while a request is active or stopping,
  - expose an `Overview` page for live request state and a `History` page for browsing recorded server events,
  - show `Database` and `Synchronization` panels on the `Overview` page with store readiness, maintenance, sync progress, active root, and live query execution mode,
  - support keyboard navigation for page switching and history browsing (`1`, `2`, `f`, arrows, `PgUp/PgDn`, `Home/End`).
  Redirected foreground-mode logs are part of the same diagnostics surface and must still report database state, synchronization state, and query execution mode for active searches when the interactive dashboard cannot render.
  `Ctrl+C` requests a clean shutdown.
- `--compact` performs offline SQLite maintenance for the selected store, acquires the single-instance guard, truncates WAL, runs `VACUUM`, records `last_checkpoint_utc` / `last_compaction_utc`, prints a before/after summary, and exits.
- `--request-compact` asks the running service to compact its live SQLite store through the named-pipe control channel, then prints refreshed DB/WAL/free-page state from `GetStatus(...)`.
- `--register` registers the current executable as the build-specific Windows service.
- `--unregister` removes the build-specific Windows service registration.
- `--pipe-name=...` can also target `--request-compact` at a non-default running service, while `--protocol-version=...`, `--max-requests=...`, and `--disconnect-after-batches=...` remain developer/test options for foreground mode.
- `--storage-root=...`, `--store-backend=...`, and `--sqlite-path=...` can be used with `--run-foreground` or `--compact`.
- the default service backend is sqlite; `--store-backend=snapshot` remains an explicit compatibility mode.
- `--storage-root=...` overrides the foreground store/storage root, which is useful for isolated self-tests and side-by-side local service runs.
- `REDSALAMANDER_SEARCH_SERVICE_PIPE` can still override the client-side pipe selection when a non-default pipe is required.

### Content search in indexed flows

In v1, indexed backends only narrow the candidate set. File contents remain client-side:
- the service or index core returns candidate paths and metadata,
- the client/plugin opens files through normal reader APIs,
- literal/regex text matching and snippet generation happen in the client process.

## Warning and Degradation Semantics

Search backends must report visible degradation through `FileSystemSearchProgress::warningFlags`, including:
- `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX`
- `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT`
- `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED`
- `FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE`
- `FILESYSTEM_SEARCH_WARNING_OVERFLOW`
- `FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED`
- `FILESYSTEM_SEARCH_WARNING_SERVICE_ROOT_REJECTED`

`FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE` means the service transport is unavailable and the host had to switch backends. `FILESYSTEM_SEARCH_WARNING_SERVICE_ROOT_REJECTED` means the service transport remained healthy but rejected this request's root and the host used another backend without arming the transport cooldown. `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX` means the backend stayed healthy but answered from the live filesystem because the index or database was not ready or not current. `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED` covers skipped traversal/candidate paths, including accepted service queries where a later per-candidate impersonation or access check cannot prove client visibility, or where a transient directory authorization open fails without a durable access-denied/path-missing answer. `FILESYSTEM_SEARCH_WARNING_OVERFLOW` also covers intentionally skipped corrupt or malformed directory/search entries when continuing with the rest of the result set is safer than failing the whole search.

The Find window must surface these warnings in status text rather than silently hiding them; rejected regex searches must show a localized rejected-regex reason rather than only the raw `E_INVALIDARG` hint. While a `service` or `local-index` search is active it must also show a second backend-status line for database readiness and synchronization progress. Host-extension service-status snapshots must apply the same database-state inference as broker-polled status so the first visible active-search status never regresses to a placeholder `search unknown` state. A ready, watching SQLite store reports no fallback before a query. Query/progress status reports the request-specific fallback reason: an unmirrored root is `StoreMissing`, a present-but-not-current root is stale/cutover-blocked, and an invalid or failed SQLite store reports the corresponding degraded reason while the request falls back to live scan.

## Persistence and Configuration

### Settings Store

Persisted Find dialog state lives under:
- `search`

This includes:
- recent roots,
- recent name patterns,
- recent content patterns,
- last-used root and pattern values,
- name/content modes,
- recursive/include/match-case/follow-symlinks toggles,
- `preferIndex`,
- `wantSnippets`,
- `maxResults`,
- result grid logical sort (`sortColumnId`, `sortDescending`) and visible column layout (`gridLayout`).

The settings store MUST sanitize persisted Find combo values at both load and save time: `recentRoots`, `recentNamePatterns`, `recentContentPatterns`, `lastRoot`, `lastNamePattern`, and `lastContentPattern` strip single-line control characters, drop entries that become empty, and de-duplicate history entries case-insensitively after sanitization. This prevents older or injected hidden shortcut characters from being preserved in configuration.

During an in-place search rerun, the Find window MUST preserve the selected result by stable full path when that result is present in the new result set. Column reorder, resize, and sort state MUST survive the clear/repopulate cycle; when the previously selected path is absent, the window may fall back to its normal first-result selection behavior.

Window placement is persisted through:
- `windows["FindFilesWindow"]`

### Built-in file plugin configuration

The built-in local file plugin persists backend preference under its plugin configuration payload:
- `searchBackendPreference`
- `searchMaxDirectoryWalkers`

Allowed `searchBackendPreference` values:
- `auto`
- `service`
- `local-index`
- `scan`

`searchMaxDirectoryWalkers` controls the bounded parallel recursive directory walk used by built-in local name-only scan searches. The effective value is clamped to `1..8`; `1` keeps the scan serial. Backend preference and walker count are independent: selecting `scan` controls backend routing, while `searchMaxDirectoryWalkers` controls only scan traversal parallelism.

## Diagnostics and Verification

Search changes must be validated through:
- Commands self-tests for the Find window, command routing, persistence, and modeless ownership,
- Commands self-tests that explicitly verify the Find window backend diagnostics line for active `service` searches, distinguish `DEGRADED_NO_INDEX` from `SERVICE_UNAVAILABLE`, and confirm the first active-search snapshot normalizes degraded backend diagnostics instead of falling back to placeholder text,
- Compare self-tests for native search, fallback parity, indexed parity, service parity, cancellation, Unicode paths, long paths, access denied, NTFS, and ReFS,
- Compare self-tests that verify corrupt or unqueryable SQLite stores still start the service, report derived pre-query database state, and fall back to live search without blocking for repair,
- Compare self-tests that verify mid-query SQLite failure restarts the request as a live scan without duplicating already emitted hits,
- Compare self-tests that verify redirected foreground-mode logs still expose database state, synchronization progress, and query execution mode when the interactive dashboard is unavailable,
- Compare self-tests that verify corrupt snapshot counts rebuild without process termination, malformed service `QueryBatch` counts fail with a fallback-eligible protocol HRESULT and fall back cleanly, and malformed USN name bounds are rejected before OOB reads,
- Compare self-test `sqlite_index_store_root_lookup_case_insensitive`, which verifies direct SQLite `InspectVolume`, `LoadVolume`, and `EnumerateVolume` find a stored root through differently-cased caller text,
- Compare self-test `search_service_sqlite_external_rotation_refreshes_without_retry`, which rotates the configured SQLite store externally mid-service-session and verifies the next query observes the refreshed generation/currentness before querying, uses SQLite for the rotated contents, and does not rely on a post-rotation `search.backend.sqlite.retry_query_ms` retry,
- Compare self-tests for Unicode/long-path parity MUST use real Unicode codepoints in the fixture names, not double-encoded mojibake, and MUST assert the fixture contains the intended codepoint before trusting the parity result,
- Compare self-tests for malformed service `QueryBatch` fallback MUST assert that fallback returns the expected local-search matches, not only that the search call succeeds,
- Compare self-test `search_service_candidate_impersonation_failure_is_incomplete_warning`, which fault-injects one service-side candidate authorization impersonation failure and verifies the query completes with unaffected candidates plus `ACCESS_DENIED_SKIPPED`,
- Compare self-test `search_service_transient_authorization_failure_is_incomplete_not_cached`, which holds a cached candidate directory with an exclusive handle to force a transient authorization `CreateFileW` failure, verifies the first query reports `ACCESS_DENIED_SKIPPED`, releases the handle, and verifies a later query in the same service session re-evaluates the directory instead of reusing a poisoned `false` cache entry,
- Compare self-tests `search_service_missing_pipe_retry_is_cancellable`, `local_search_service_root_rejection_does_not_arm_transport_cooldown`, and `search_service_transient_parent_failure_is_cached_and_warns_batch_and_completion`, which verify cancellable absent-service retries, request-scoped root rejection without global transport cooldown, per-query transient-parent suppression, and warning parity between streamed batches and completion,
- Compare self-tests `local_search_sqlite_generation_probe_skips_steady_state_and_detects_bump`, `local_search_junction_alias_hydration_indexes_alias_only_descendant_once`, and `sqlite_maintenance_busy_checkpoint_and_callback_skip_observability`, which verify zero steady-state generation-probe opens after warm-up with next-query rotation detection, coverage-only alias hydration without duplicate emission, and truthful checkpoint/emitted-row observability,
- Compare self-tests for regex rejection MUST cover both name and content regex routes, including known unsafe alternation patterns plus configured pattern-length and group-depth limits,
- Compare self-test `search_source_allocation_and_folding_guard`, which verifies QueryBatch count preflight before reserve, no-callback broker buffering ceilings before accumulation, corrupt snapshot entry-count preflight before reserve, invariant search-fold callsites, pre-folded prefix queries for in-pane incremental search, absence of broad `catch (...)`, and fatal `std::bad_alloc` handling on search paths,
- Compare and Commands adversarial fault-injection coverage MUST include stalled broker queries, malformed `QueryBatch` direct decode and fallback, corrupt snapshot rebuild, journal replay invalidation rebuild, native and fallback symlink-loop traversal bounds, and invalid-regex or cancelled Intersect preservation,
- Compare self-test `search_low_hardening_smoke`, which verifies snippet boundaries preserve UTF-16 surrogate pairs, `FILE_FULL_DIR_INFO` name bounds are checked before reads, followed-symlink visit keys roll back on physical-identity duplicate rejection and identity-probe failure, malformed parallel entries are skipped with warnings instead of failing the directory walk or abandoning the chunk, rejected-regex warnings are surfaced in Find status, empty plugin-factory spans are guarded, ViewerWeb internal HTML head/CSP literals stay hoisted, snapshot saves use flushed atomic sibling replacement, SQLite bootstrap does not run full `VACUUM` inline, SQLite maintenance runs `quick_check` before full `VACUUM`, automatic SQLite maintenance owns legacy auto-vacuum rewrites, hardlink alias non-exhaustiveness is reported through query stats, and `ReplaceVolume` performs one metadata update for an existing volume,
- Compare self-tests `search_service_rejects_device_root_and_continues` and `search_service_slow_partial_client_does_not_block_next_client`, which verify service-side root rejection for device namespaces, client-scoped failure without service termination, bounded server frame reads, and slow/partial-client disconnect-and-continue behavior,
- Compare self-test `search_service_rebuild_deleted_root_purges_index`, which verifies that `RequestRebuild(root)` succeeds after an indexed root directory is deleted and purges stale indexed candidates before a later query of the recreated root,
- Compare self-test `search_service_sqlite_legacy_auto_vacuum_queues_idle_maintenance`, which verifies a legacy non-incremental SQLite store queues idle maintenance and rewrites to incremental auto-vacuum without requiring WAL or freelist pressure; its first exact-pipe status response establishes both foreground-service readiness and the queued-maintenance state, and completion polling tolerates the bounded pipe-unavailable interval while synchronous maintenance owns the service loop,
- Compare foreground-service coverage must establish readiness through a successful status response whose returned pipe exactly matches the intended isolated pipe, request explicit graceful shutdown instead of guessing a request count, and prove process exit inside a bounded two-second owner-controlled window,
- Compare SQLite status coverage must distinguish `ImportedLegacySnapshot` from `CurrentnessUnproven`: both block cutover, but only the former increments legacy-import migration state; pre-query store state remains syncing/idle rather than ready/watching for either unready volume state,
- that status coverage must include a mixed Ready + `CurrentnessUnproven` store and prove the unready volume keeps aggregate cutover blocked without hiding the Ready volume from the indexed-volume summary,
- warm-up status coverage must compare service readiness with the states actually persisted for every warmed volume rather than assuming that every machine can prove USN currentness; when the aggregate store remains unready, its query must take the live-scan fallback, while capability-based direct-SQLite assertions remain conditional on a readable live journal cursor,
- Commands self-test `search_local_plugin_invalid_regex_reports_single_completion`, which verifies built-in local syntax-invalid and safety-rejected regex searches return `E_INVALIDARG`, emit no matches, and report the single completed progress payload described above,
- Commands self-test `search_local_plugin_parallel_cancel_fanin`, which verifies built-in local parallel name-only scan finds flat, broad, and deep-tree matches, serializes fan-in callbacks, and honors `FileSystemSearchShouldCancel` after a ready match batch starts draining,
- Commands self-test `filesystem_local_watch_unwatch_drains_inflight_callback`, which verifies built-in local watch teardown drains a blocked in-flight callback before `UnwatchDirectory(...)` returns,
- packaged Release validation for MSI and MSIX when service packaging changes.

Machine-dependent coverage remains part of the suite and must skip with a reason when preconditions are absent. Example:
- the ReFS indexed probe remains declared and records `skipped` when no fixed ReFS volume exists.

## Normative Summary

- `cmd/pane/find` is implemented and opens the modeless Find window.
- The Find window is modeless, independently instanced, themed, and worker-threaded.
- `IFileSystemSearch` is the native plugin contract; the host must provide scan fallback when native search is unavailable.
- The built-in local plugin prefers `service -> local-index -> scan`.
- Debug and Release service builds use different default ProgramData roots so they can run side by side without sharing SQLite state.
- SQLite bootstrap or inspection failures must be fail-open at service startup so status polling and degraded live-search queries still work.
- SQLite-backed service queries must choose the fastest no-wait path per request and degrade to live filesystem search when currentness is not proven.
- SQLite-backed service queries must validate cached store generation before direct SQLite use, and retry/dedupe stale-store fallbacks without duplicating hits that already reached the client.
- Service request roots are a trust boundary: query/rebuild roots must be canonicalized, device/UNC/relative roots rejected for the local service backend, and accepted roots authorized under the named-pipe client's impersonated identity before service authority touches them. Rebuild/invalidate requests may authorize a missing leaf through the deepest existing ancestor, but the purge target remains the original normalized requested root.
- Search service pipe I/O must be bounded, overlapped, cancellable on shutdown, and tolerant of malformed/slow clients by disconnecting only the failed client and continuing to accept later clients.
- Service and host status polling must derive the initial database state from persistent-store inspection and startup warm-up progress even before the first query updates repository runtime state.
- Persistent-store readiness is true only when every indexed volume is in the `Ready` state; legacy-migration counts and query-readiness counts are distinct contracts.
- Foreground service ownership uses exact-pipe status readiness plus an explicit graceful-shutdown request. Request-count budgets are not a lifecycle or readiness mechanism, and SCM-hosted services reject foreground shutdown requests.
- Host-extension relays of service status must preserve that same inferred database state so Find-window diagnostics stay consistent before query-specific progress arrives.
- Indexed backends return names, paths, and metadata; content matching remains client-side in v1.
- Regex syntax errors and shared safety rejects must complete with one `REGEX_REJECTED` progress payload, no matches, and `E_INVALIDARG`; oversized decoded regex content must surface `WARNING_OVERFLOW` instead of looking like a clean no-match.
- The Find window must render `REGEX_REJECTED` as a localized status warning instead of leaving users with only the raw `E_INVALIDARG` code.
- Built-in local name-only scan searches may parallelize recursive traversal through `searchMaxDirectoryWalkers`, bounded to `1..8`.
- Unchecked Find indexed-backend preference forces built-in local searches to the scan backend instead of allowing service/local-index selection through `auto`.
- Parallel search fan-in and indexed/broker result streams must poll cancellation during match emission or between service batches, serialize callbacks, and reach a worker quiet point before returning cancellation/failure.
- Built-in local parallel search workers must be coordinator-owned `PTP_WORK`/RAII work items or an equivalent cleanup-group ownership model. Raw `TrySubmitThreadpoolCallback` callbacks must not receive pointers to stack-owned scan state, condition variables, chunk vectors, or module pins unless the coordinator has a proven wait that outlives every possible callback access.
- Search snippets must not split UTF-16 surrogate pairs at the snippet start or end boundary.
- Followed-symlink traversal must de-duplicate physical directory identities without leaving stale logical visit keys for rejected duplicate identities, and must skip followed local directories whose physical identity cannot be probed.
- External/corruptible search counts and byte lengths must be validated before allocation; protocol corruption must use a failing protocol HRESULT, configured accumulation ceilings must fail before unbounded retention, and true OOM remains fatal.
- USN and directory-information fixed-offset name fields must validate offset/length bounds against the current record or entry before reading the name.
- SQLite bootstrap must not run full `VACUUM` inline; full rewrites belong to manual or automatic maintenance, automatic maintenance must persist incremental auto-vacuum mode on upgraded legacy stores, volume replacement should update volume metadata once per transaction, and SQLite store generation must change on committed store content/currentness mutations so query-time cache validation can detect external rotation.
- Every visible degradation must be surfaced through progress warnings and reflected in the UI together with database or synchronization status when available.
- Redirected foreground-mode logs are part of the service diagnostics surface and must carry the same database, synchronization, and query-execution-state information that the interactive dashboard shows for active searches.
