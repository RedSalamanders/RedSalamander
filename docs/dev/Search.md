# Search Subsystem (developer)

This page is the developer map for RedSalamander search: the modeless Find window, its worker-threaded session controller, the host-side backend reporting, and the out-of-process SQLite-backed search service. For the user-facing dialog see [../FindFiles.md](../FindFiles.md); for the higher-level subsystem walkthrough see the "Find Files, Search Backends & the Search Service" section of [../DeveloperGuide.md](../DeveloperGuide.md). The authoritative behavior spec is `Specs/Core/Core_Search.md`.

## Layout

Search is a host-owned feature: one Find experience for every `IFileSystem` plugin, a fast indexed/service backend preferred for local roots, and a live-scan correctness baseline that always works.

| File / type | Role |
|---|---|
| `RedSalamander/FindFilesWindow.{h,cpp}` — `FindFilesWindow` | Modeless DxUI window: UI, result grid, set algebra, status/backend lines |
| `RedSalamander/FindFilesWindow.cpp` — `SearchSessionController` | Owns the `std::jthread` worker, cancellation, completion delivery |
| `RedSalamander/SearchFallbackEngine.{h,cpp}` — `SearchFallbackEngine::Execute` | Generic scan baseline over `IFileSystem::ReadDirectoryInfo` |
| `Common/SearchServiceBroker.{h,cpp}` | Named-pipe client (`GetStatus`/`Query`/`RequestRebuild`/`RequestCompact`) plus `RunServer` |
| `Common/LocalSearchIndexCore.{h,cpp}` — `Repository` | Shared index core: probe, query planning, journal replay |
| `Common/SqliteIndexStore.{h,cpp}` | `index-v2.sqlite3` schema-v2 read/write |
| `RedSalamanderSearchService/Main.cpp` | Service host, CLI, and `--run-foreground` FTXUI dashboard |
| `Plugins/FileSystem/FileSystem.Search.cpp` | Built-in local `file` plugin's `IFileSystemSearch` and its backend routing |

## Find window and the session controller

`cmd/pane/find` (default shortcuts `Alt+F7` and `Ctrl+F`) calls `ShowFindFilesWindow(owner, settings, theme, context)` (`RedSalamander/FindFilesWindow.h`). The `FindFilesPaneContext` snapshots the target pane's `IFileSystem`, `pluginId`, `pluginShortId`, `instanceContext`, and `rootPluginPath`. Each invocation creates an independent modeless `FindFilesWindow` (a DxUI host control implementing `IDxGridDelegate`); multiple Find windows can be open at once, each with its own result set.

`BeginSearch(operation, textOverride)` builds a `SearchRequest` from the combos and checkboxes (root path, name/content patterns and modes, flags, `maxResults`, content limits) and calls `SearchSessionController::Start(owner, request)`. The controller owns:

- the worker thread (`std::jthread _worker`),
- cancellation (`std::atomic<bool> _cancelRequested`),
- active/idle state (`_active`, `_idleCv`), and a `_uiSettled` flag fed by `NotifyUiSettled()`.

`Start` spawns `Run(request)` on the worker. `Run` fills a `FileSystemSearchQuery`, `QueryInterface`s the plugin for `IFileSystemSearch`, and either calls `Search(...)` or, when the plugin has no search interface (or `Search` reports unsupported and the search was not cancelled), falls back to `SearchFallbackEngine::Execute(...)`. The fallback traverses through `IFileSystem::ReadDirectoryInfo` and reads content through `IFileSystemIO::CreateFileReader` when available; if no readable stream API exists, content search degrades and reports `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT`.

## FileSystemSearchBackend (and why there is no host `SelectSearchBackend`)

`FileSystemSearchBackend` (`Common/PlugInterfaces/FileSystem.h`) is the enum a backend reports through `FileSystemSearchProgress::backend`:

| Value | Meaning |
|---|---|
| `FILESYSTEM_SEARCH_BACKEND_UNKNOWN` (0) | Not yet determined |
| `FILESYSTEM_SEARCH_BACKEND_SCAN` (1) | Live directory traversal |
| `FILESYSTEM_SEARCH_BACKEND_INDEX` (2) | In-process `LocalSearchIndexCore` index |
| `FILESYSTEM_SEARCH_BACKEND_SERVICE` (3) | Out-of-process search service |

The host does **not** itself pick a backend. There is no `SelectSearchBackend` symbol in `FindFilesWindow.cpp` / `SearchSessionController` — the host only chooses *native search vs. scan fallback* and then reports whatever `FileSystemSearchBackend` the active backend hands back in progress payloads (`FindFilesWindow` maps it to a status string via `BackendStringId(...)`).

Backend *selection* (`service` -> `local-index` -> `scan`) is internal to the built-in local `file` plugin. In `Plugins/FileSystem/FileSystem.Search.cpp`, the plugin's private `SelectSearchBackend(preference, flags, support, preferServiceBackend)` routes on a `FileSystemSearchBackendPreference` (`Auto`/`Service`/`LocalIndex`/`Scan`, persisted as `searchBackendPreference`). `FILESYSTEM_SEARCH_FORCE_SCAN` always wins and forces `SCAN`; the Find window sets that flag when "Prefer indexed backend" is unchecked, so unchecked means "scan now", not "let `auto` choose".

For the built-in local plugin the host enables internal extensions: `Run` sets `query.reserved = FILESYSTEM_SEARCH_HOST_EXTENSIONS_V1` ("RSF1") and passes a `FileSystemSearchHostExtensions` cookie after seeding `SearchServiceBroker::GetStatus(...)` so the first active-search status line already reflects service/database state.

## Worker → UI message IDs

`IFileSystemSearch::Search(...)` is synchronous and runs on the worker, so results never touch the UI thread directly. The worker batches `FindResultRecord`s and `PostMessagePayload`s them; `FindFilesWindow::WindowProc` consumes them. The IDs are defined in `Common/WindowMessages.h`:

| Message | Value | Handler | Payload |
|---|---|---|---|
| `WndMsg::kFindSearchResults` | `WM_APP + 0x527` | `OnSearchResults` | `FindSearchResultsPayload` |
| `WndMsg::kFindSearchProgress` | `WM_APP + 0x528` | `OnSearchProgress` | `FindSearchProgressPayload` |
| `WndMsg::kFindSearchComplete` | `WM_APP + 0x529` | `OnSearchComplete` | `FindSearchCompletePayload` |
| `WndMsg::kFindSearchDeferredRefresh` | `WM_APP + 0x52A` | `ApplyPendingResultsRefresh` | none |
| `WndMsg::kFindShowActionMenu` | `WM_APP + 0x52D` | shows the Find split-button menu | none |

Receivers reclaim the payload with `TakeMessagePayload<T>(lParam)`. Result batches flush by capacity or age (never on a low-priority timer) so a `*` flood produces a first visible batch promptly. Cancellation is checked at traversal boundaries and while draining ready batches; a cancelled search must leave no worker/search state behind.

## Backend precedence and degradation

For the built-in local `file` plugin the precedence is `service` → `local-index` → `scan`; `auto` walks that order. Supported indexed roots are NTFS and ReFS — FAT/exFAT, UNC, WSL-backed, and other unsupported roots fall back to scan. Visible degradation flows through `FileSystemSearchProgress::warningFlags`:

- `FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE` — the service transport was unavailable and the host switched backends.
- `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX` — the backend stayed healthy but answered from the live filesystem (index/DB not ready or not current).
- `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT`, `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED`, `FILESYSTEM_SEARCH_WARNING_OVERFLOW`, `FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED`.

While a `service` or `local-index` search is active, the Find window shows a second backend-status line built from `LocalSearchIndexCore` runtime state: `StoreState`, `SyncPhase`, `QueryExecutionMode`, and `FallbackReason`.

## SearchServiceBroker and the named-pipe protocol (v3)

`Common/SearchServiceBroker.h` defines both the in-process client API and the server entry point. Build-specific identities:

| Build | Service name | Pipe |
|---|---|---|
| Debug | `RedSalamanderSearchService.Debug` | `\\.\pipe\RedSalamander.SearchService.Debug.v3` |
| Release | `RedSalamanderSearchService` | `\\.\pipe\RedSalamander.SearchService.v3` |

`kProtocolVersion = 3`. `REDSALAMANDER_SEARCH_SERVICE_PIPE` overrides the client-side pipe selection.

Client entry points: `GetStatus(ServiceStatus&)`, `Query(QueryRequest, progressCallback, cookie, cancelCheck, cancelCookie, outCandidates, outStats, candidateBatchCallback, candidateBatchCookie)`, `RequestRebuild(rootPath)`, `RequestCompact()`. The server runs through `RunServer(options, stopEvent, outResult)`.

### Framing

Every message is a fixed `FrameHeader` followed by a variable payload, written/read with `SendFrame` / `ReceiveFrame`:

```text
struct FrameHeader {
    uint32_t magic;            // 0x53535252 ("RRSS")
    uint32_t protocolVersion;  // 3
    uint32_t messageType;      // MessageType
    uint32_t payloadBytes;     // <= 16 MiB (kMaxFrameBytes)
};
```

A frame with a wrong magic is rejected as `RPC_S_PROTOCOL_ERROR`; an oversized `payloadBytes` is rejected as `ERROR_BUFFER_OVERFLOW`. The client also rejects a response whose `protocolVersion` differs from `kProtocolVersion`.

`MessageType` values:

| Type | Value | Direction | Notes |
|---|---|---|---|
| `StatusRequest` | 1 | client → server | empty payload; answered by `StatusResponse` |
| `StatusResponse` | 2 | server → client | base + extended/maintenance/discovery/warmup/runtime sub-payloads |
| `QueryRequest` | 3 | client → server | `QueryRequestPayload` + UTF-16 root path + name pattern |
| `QueryProgress` | 4 | server → client | streamed `ProgressPayload` (+ `ProgressRuntimePayload`) |
| `QueryBatch` | 5 | server → client | streamed `CandidateBatchHeader` + N `CandidateEntryHeader` + strings |
| `QueryComplete` | 6 | server → client | terminal `QueryCompletePayload` (+ runtime payload + flags) |
| `RebuildRequest` | 7 | client → server | UTF-16 root path; answered by `Ack` or `Error` |
| `Ack` | 8 | server → client | success for rebuild/compact control requests |
| `Error` | 9 | server → client | carries an `HRESULT` result |
| `CompactRequest` | 10 | client → server | empty payload; answered by `Ack` or `Error` |

The server dispatches inbound frames in `HandleStatusRequest` / `HandleQueryRequest` / `HandleRebuildRequest` / `HandleCompactRequest`; it never accepts a *response* type inbound (those map to `RPC_S_PROTOCOL_ERROR`). These map onto the `ServerRequestType` reported in server events (`SEARCH_SERVICE_SERVER_REQUEST_STATUS/QUERY/REBUILD/COMPACT`).

### Streaming and cancellation

A query is a streamed exchange: the client sends one `QueryRequest`, then loops receiving frames until a terminal `QueryComplete` or `Error`. While looping it interleaves `QueryProgress` (forwarded to `progressCallback`) and `QueryBatch` candidate batches (forwarded to `candidateBatchCallback`, with `consumedCount` so the caller can apply a `maxResults` cutoff). A `QueryBatch` is `CandidateBatchHeader { count }` followed by `count` `CandidateEntryHeader` records, each followed by its UTF-16 `fullPath` and `displayName`; the server buffers entries and flushes when the next entry would exceed `kMaxFrameBytes`.

Cancellation is cooperative and client-driven: at the top of each receive iteration the client calls `cancelCheck(cancelCookie)`. A failing HRESULT (the controller's cancel flag mapped to a cancel HRESULT) makes the client return immediately and drop the pipe; the disconnect signals the server to abandon the in-flight query. There is no separate "cancel" frame.

`QueryComplete` carries a `QueryCompletePayload` (`result` HRESULT, file-system kind, counts, journal id / next USN, durations) plus a `flags` bitfield (`QUERY_COMPLETE_FLAG_USED_SQLITE_STORE`, `..._USED_NAME_PREFILTER`, `..._SQLITE_CUTOVER_BLOCKED`, `..._JOURNAL_REPLAY_APPLIED`, `..._USED_TRAVERSAL_SEED`, ...) and a runtime payload that fixes the final `QueryExecutionMode` / `FallbackReason`. If a direct SQLite query fails after emitting early hits, the request is restarted on the live filesystem and duplicate hits are suppressed rather than surfacing a hard failure.

### Service CLI (`RedSalamanderSearchService.exe`)

`Main.cpp` maps startup actions to flags: `--run-foreground` (terminal + full-screen FTXUI dashboard, `Ctrl+C` for clean shutdown), `--compact` (offline checkpoint/VACUUM and exit), `--request-compact` (ask the running service to compact over the pipe), `--register` / `--unregister` (SCM registration, usually elevated). Store/transport overrides: `--pipe-name=`, `--storage-root=`, `--store-backend=` (`sqlite` default, `snapshot` compatibility), `--sqlite-path=`. Default storage roots are `%ProgramData%\RedSalamander\SearchIndex` (Release) and `...\SearchIndex.Debug` (Debug); the default SQLite store is `index-v2.sqlite3` under that root. Debug and Release roots stay behaviorally isolated so the two builds can run side by side.

## LocalSearchIndexCore::Repository

`Common/LocalSearchIndexCore.h` is the shared index core used by both the built-in `file` plugin and the service. `Repository` exposes `ProbePath`, `Query`, `Enumerate` / `EnumerateNoWait`, `PrimeRoot`, `EnsureReadyForRoot`, `InvalidateRoot`, `GetStatus`, and `CollectCachedRoots`. `RepositoryOptions` selects a `PersistentStoreKind` (`SnapshotBinary` = 1, `Sqlite` = 2), a snapshot root and/or SQLite path, and `sqliteAuthoritative`. The runtime state types reported through progress and status:

| Enum | Values |
|---|---|
| `FileSystemKind` | `Unsupported` (0), `Ntfs` (1), `Refs` (2) |
| `StoreState` | `Unknown`, `Ready`, `Syncing`, `Recovering`, `Invalid`, `Maintenance` |
| `SyncPhase` | `Idle`, `Discovering`, `Loading`, `Replaying`, `Mirroring`, `Watching` |
| `QueryExecutionMode` | `Unknown`, `Sqlite`, `InMemoryIndex`, `LiveScanFallback` |
| `FallbackReason` | `None`, `StoreMissing`, `StoreInvalid`, `StoreStale`, `CutoverBlocked`, `WarmupRunning`, `SqliteFailure` |

The core is responsible for path normalization, query planning (`QueryPlan`), journal replay, traversal-seed rebuilds when direct journal/MFT access is unavailable, and (for the service) SQLite persistence. In v1 indexed backends only narrow the candidate set — they return paths and metadata (`Candidate`), and literal/regex content matching and snippet generation happen client-side.

## SqliteIndexStore (schema v2)

`Common/SqliteIndexStore.h` reads and writes the SQLite store. `kSchemaVersion = 2`. The store is the authoritative persisted current state plus a journal cursor for service-managed roots; it persists filesystem metadata and root journal state but never file contents, snippets, or query history.

Tables (per `Specs/Core/Core_Search.md`):

- `meta(key, value)` — `schema_version`, `store_kind`, `last_checkpoint_utc`, `last_compaction_utc`.
- `volumes(volume_id, root_path, fs_kind, journal_id, next_usn, state, entry_count, last_seed_utc, last_replay_utc, last_error_hr)` — per-root indexing state.
- `entries(volume_id, file_id_low, file_id_high, parent_id_low, parent_id_high, full_path, full_path_folded, name, name_folded, extension_folded, attributes, is_dir, size_bytes, write_time_100ns, creation_time_100ns, last_access_time_100ns, change_time_100ns, allocation_size)` — indexed metadata, with case-folded columns for invariant Windows-style matching.

Volume `state` codes (`SqliteIndexStore.h`):

| Constant | Value | Meaning |
|---|---|---|
| `kVolumeStateReady` | 1 | Current and queryable directly |
| `kVolumeStateImportedLegacySnapshot` | 2 | Imported from a legacy binary snapshot |
| `kVolumeStateCurrentnessUnproven` | 3 | Seeded by traversal (e.g. non-elevated, no journal); cutover stays degraded until currentness is proven |

Key functions: `EnsureBootstrap`, `InspectStore`, `InspectVolume`, `LoadVolume`, `ReplaceVolume`, `ApplyJournalDelta`, `DeleteVolume`, `EnumerateVolume`, `RunManualMaintenance` (offline VACUUM), and `RunAutomaticMaintenance` (policy-driven checkpoint + incremental vacuum). `EnsureBootstrap` is fail-open: bootstrap or inspection failures still let the service start, report derived pre-query state, and answer from live scan.

## DirectoryInfoCache is not the Find traversal path

`DirectoryInfoCache` (`RedSalamander/DirectoryInfoCache.h`, spec `../../Specs/Core/Core_DirectoryInfoCache.md`) is the app-global cache of folder enumeration snapshots (`IFilesInformation`) and the host-side mutation router that keeps visible panes coherent. It is **not** part of Find. Find traversal goes straight through `IFileSystem::ReadDirectoryInfo` in `SearchFallbackEngine` (or through the indexed/service backends) and deliberately bypasses the cache — searching a large tree must not pollute or thrash the pane snapshot cache, and the cache's borrow/pin/watch model is tuned for visible panes, not bulk traversal. When verifying Find behavior, look at `SearchSessionController` and the backends, not at `DirectoryInfoCache`.

## Diagnostics

Search work prefers the existing `find.ui.*` / `find.session.*` perf metrics. Relevant self-tests live under `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp` (e.g. `search_local_plugin_invalid_regex_reports_single_completion`, `search_local_plugin_parallel_cancel_fanin`) and the search/index compare cases under `RedSalamander/SelfTest/CompareDirectories/`. See `Specs/Core/Core_Search.md` for the full performance-validation and verification contract.
