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

- `cmd/pane/find` opens or activates the modeless `Find Files and Directories` window.
- The command targets the focused pane when focus is inside a pane; otherwise it targets the active pane.
- The default shortcuts are `Alt+F7` and `Ctrl+F`.
- If the window already exists, the host must reuse it and refresh its pane context instead of opening duplicates.

### Default scope

The initial search scope is derived from the target pane:
- active filesystem plugin,
- active instance context,
- current pane path as `rootPath`.

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
- Name and content regex patterns MUST be validated before traversal or content reads begin. Syntax errors and safety rejects (for example nested unbounded repetition or excessive nesting/length) MUST fail with `E_INVALIDARG`, emit no matches, and report exactly one forced `FILESYSTEM_SEARCH_PHASE_COMPLETED` progress payload with `FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED` and `statusHint == E_INVALIDARG`.
- A rejected regex search MUST leave no active worker/search state behind. For `IFileSystemSearch::Search(...)`, callbacks remain per-call and MUST NOT be invoked after the method returns.

## Result Model

### Set operations

The Find window supports four result-set operations:
- `Find`: replace the current result set.
- `Append`: union the new result set into the current result set.
- `Intersect`: keep only items present in both sets.
- `Subtract`: remove newly found items from the current result set.

The visible Find action is a split button: its primary action runs `Find`, and its menu exposes `Find`, `Intersect`, `Subtract`, and `Append`. Set operations other than `Find` MUST remain disabled while the current result set is empty, and MUST become available once at least one result is present.

### Canonical result identity

Result identity is frozen as:
- `pluginId + normalized instanceContext + normalized fullPath`

This identity is used by the Find window for de-duplication and set algebra across native, indexed, service, and fallback backends.

### Result actions

- Double-clicking a file opens it through the existing host open/view flow.
- Double-clicking a directory navigates the pane to that directory.
- The secondary parent action navigates to the parent directory and focuses the matched item when possible.
- Recursive results MUST display the containing subfolder relative to the search root in the Path column, including when a backend reports only a leaf relative path but provides a full path under the requested root.
- One-line result rows MUST use the shared grid density metrics and shrink when compact mode is active; snippet rows may remain taller to preserve preview readability.

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
- Cancellation must be checked at traversal boundaries and during chunked content reads.
- Built-in local parallel name-only scan MUST also check cancellation while draining completed worker result batches. A large directory that produces many ready matches MUST NOT defer `FileSystemSearchShouldCancel` until the whole batch has been emitted.
- `IFileSystemSearch` callbacks for a single `Search(...)` call MUST be serialized even when parallel scan workers finish concurrently. The coordinator may abandon queued completed batches after a terminal cancellation/failure once all workers have reached the quiet point.
- Parallel recursive scan MUST avoid starvation: broad sibling directories and deep descendant chains both have to make progress unless the search is cancelled or a terminal error occurs.

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

### Generic host scan fallback

The host scan fallback is the correctness baseline for all plugins. It uses:
- `IFileSystem::ReadDirectoryInfo` for traversal,
- `IFileSystemIO::CreateFileReader` for content scanning when available.

Fallback behavior:
- If a plugin does not expose `IFileSystemSearch`, the host may still provide name search through scan fallback.
- Content fallback additionally requires `IFileSystemIO`.
- If no readable stream API exists, content search must degrade cleanly and surface `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT`.

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
  - `meta(key, value)` for store metadata such as `schema_version`, `store_kind`, `last_checkpoint_utc`, and `last_compaction_utc`,
  - `volumes(volume_id, root_path, fs_kind, journal_id, next_usn, state, entry_count, last_seed_utc, last_replay_utc, last_error_hr)` for per-root indexing state,
  - `entries(volume_id, file_id_low, file_id_high, parent_id_low, parent_id_high, full_path, full_path_folded, name, name_folded, extension_folded, attributes, is_dir, size_bytes, write_time_100ns, creation_time_100ns, last_access_time_100ns, change_time_100ns, allocation_size)` for indexed filesystem metadata,
- persist indexed filesystem metadata, root journal state, and maintenance timestamps, but not file contents, query history, snippets, or transient in-memory synchronization/query-progress state,
- treat the SQLite store as the authoritative persisted current state plus journal cursor for service-managed roots,
- accept status and query requests before startup warm-up, rebuild, repair, or SQLite mirroring completes,
- choose the fastest no-wait query path per request:
  - direct SQLite only when the configured store is valid and current for the requested root,
  - live filesystem scan fallback otherwise,
- never block the request thread on `EnsureReady`, rebuild, repair, or SQLite mirroring,
- keep service request-thread routing strictly two-path in sqlite-authoritative mode:
  - direct SQLite only when currentness is provably up to date for the requested root,
  - live filesystem scan fallback otherwise while background warm-up or repair continues,
- when NTFS journal/MFT access is unavailable because the service is not elevated, reuse the same traversal seed strategy used for ReFS, persist that traversal seed into SQLite as currentness-unproven, and keep query cutover degraded until a later run can prove currentness from the live journal,
- keep file content out of the service protocol in v1,
- report missing, invalid, stale, and cutover-blocked SQLite state by degrading to live filesystem search with `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX`,
- if a direct SQLite query fails after already emitting early hits, restart the request as a live filesystem scan and suppress duplicates instead of surfacing a hard failure,
- use invariant Windows-style case folding for indexed and fallback path or name matching so `A.txt` and `a.txt` match the same entry while composed and decomposed Unicode names remain distinct,
- fall back safely to `local-index` or `scan` when the service is absent, mismatched, disconnected, or unavailable.

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

`FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE` means the service transport is unavailable and the host had to switch backends. `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX` means the backend stayed healthy but answered from the live filesystem because the index or database was not ready or not current.

The Find window must surface these warnings in status text rather than silently hiding them, and while a `service` or `local-index` search is active it must also show a second backend-status line for database readiness and synchronization progress. Host-extension service-status snapshots must apply the same degraded-mode inference as broker-polled status so the first visible active-search status never regresses to a placeholder `search unknown` state when warm-up or cutover state already implies live-scan fallback.

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
- Commands self-test `search_local_plugin_invalid_regex_reports_single_completion`, which verifies built-in local syntax-invalid and safety-rejected regex searches return `E_INVALIDARG`, emit no matches, and report the single completed progress payload described above,
- Commands self-test `search_local_plugin_parallel_cancel_fanin`, which verifies built-in local parallel name-only scan finds flat, broad, and deep-tree matches, serializes fan-in callbacks, and honors `FileSystemSearchShouldCancel` after a ready match batch starts draining,
- Commands self-test `filesystem_local_watch_unwatch_drains_inflight_callback`, which verifies built-in local watch teardown drains a blocked in-flight callback before `UnwatchDirectory(...)` returns,
- packaged Release validation for MSI and MSIX when service packaging changes.

Machine-dependent coverage remains part of the suite and must skip with a reason when preconditions are absent. Example:
- the ReFS indexed probe remains declared and records `skipped` when no fixed ReFS volume exists.

## Normative Summary

- `cmd/pane/find` is implemented and opens the modeless Find window.
- The Find window is modeless, independent, themed, and worker-threaded.
- `IFileSystemSearch` is the native plugin contract; the host must provide scan fallback when native search is unavailable.
- The built-in local plugin prefers `service -> local-index -> scan`.
- Debug and Release service builds use different default ProgramData roots so they can run side by side without sharing SQLite state.
- SQLite bootstrap or inspection failures must be fail-open at service startup so status polling and degraded live-search queries still work.
- SQLite-backed service queries must choose the fastest no-wait path per request and degrade to live filesystem search when currentness is not proven.
- SQLite-backed service queries that fail after emitting partial results must restart on the live filesystem path without duplicating hits that already reached the client.
- Service and host status polling must derive the initial database state from persistent-store inspection and startup warm-up progress even before the first query updates repository runtime state.
- Host-extension relays of service status must preserve that same inferred degraded state so Find-window diagnostics stay consistent before query-specific progress arrives.
- Indexed backends return names, paths, and metadata; content matching remains client-side in v1.
- Regex syntax errors and safety rejects must complete with one `REGEX_REJECTED` progress payload, no matches, and `E_INVALIDARG`.
- Built-in local name-only scan searches may parallelize recursive traversal through `searchMaxDirectoryWalkers`, bounded to `1..8`.
- Parallel search fan-in must poll cancellation during match emission, serialize callbacks, and reach a worker quiet point before returning cancellation/failure.
- Every visible degradation must be surfaced through progress warnings and reflected in the UI together with database or synchronization status when available.
- Redirected foreground-mode logs are part of the service diagnostics surface and must carry the same database, synchronization, and query-execution-state information that the interactive dashboard shows for active searches.
