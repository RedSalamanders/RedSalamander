# Plugin Host Model & Cross-FS Bridge

This page is the developer reference for how RedSalamander discovers, loads, and drives COM-style file-system plugins, how plugins call host services back across thread boundaries, and how a copy/move between two different providers is performed by the host-owned streaming bridge.

For the higher-level overview and the user-facing plugin list, see the [Plugin Host Model & the Cross-File-System Bridge](../DeveloperGuide.md#plugin-host-model--the-cross-file-system-bridge) section of the Developer Guide and [Plugins.md](../Plugins.md). The normative ABI contract lives in `Specs/Plugins/Plugins_VirtualFileSystem.md` and `Specs/Plugins/Plugins_PluginAPI.md`; the copy/move conflict contract lives in `Specs/FileSystem/FileSystem_FileOperations.md`.

## Big picture

A pane never calls Win32 file APIs directly. Each pane holds an `IFileSystem` instance produced by a plugin DLL. Three subsystems make this work:

| Subsystem | File / type | Role |
|---|---|---|
| Plugin manager | `RedSalamander/FileSystemPluginManager.cpp` | Discover, load, activate, unload, and refresh provider DLLs |
| Host services | `RedSalamander/HostServices.cpp` (`HostServices`) | Implement `IHost` + sibling interfaces; marshal UI work to the window thread |
| Cross-FS bridge | `RedSalamander/FolderWindow.FileOperations.State.cpp` (`CrossFileSystemBridge`) | Stream bytes reader -> temp -> promote between two unrelated providers |
| Capability gate | `RedSalamander/FolderWindow.FileOperations.cpp` (`CanCrossFileSystemCopyMove`) | Parse provider `GetCapabilities` JSON and decide whether a cross-FS transfer is allowed |
| External actions | `RedSalamander/FileActionResolver.cpp`, `FileActionLauncher.cpp` | Resolve and launch external view/edit programs |

The plugin-facing interfaces are defined in `Common/PlugInterfaces/` (`Factory.h`, `Host.h`, `FileSystem.h`, `Informations.h`).

## FileSystemPluginManager lifecycle

`FileSystemPluginManager` is a process singleton (`GetInstance()`). Plugin DLLs export the C entry points declared in `Common/PlugInterfaces/Factory.h`:

| Export | Purpose |
|---|---|
| `RedSalamanderEnumeratePlugins` | List `PluginMetaData` rows; one DLL may advertise several logical plugins |
| `RedSalamanderCreate` | Instantiate the requested plugin id, given an `IHost*` and `FactoryOptions` |
| `RedSalamanderGetConfigurationSchema` | Optional static schema fetch without a live instance |
| `RedSalamanderPluginShutdown` | Optional module-level quiet point (idempotent, non-throwing) |
| `RedSalamanderPluginRetainModuleUntilProcessExit` | Optional: keep the DLL mapped for OS process teardown |

### Discover and load

`Discover()` builds candidates from three origins captured in `PluginEntry::origin`:

- `PluginOrigin::Embedded` — built-in DLLs shipped next to the executable.
- `PluginOrigin::Optional` — every DLL under the optional `Plugins\` directory.
- `PluginOrigin::Custom` — paths the user added via `AddCustomPluginPath` (persisted in `settings.plugins.customPluginPaths`).

For each DLL, the manager calls `RedSalamanderEnumeratePlugins` to read `PluginMetaData` rows. A `PluginEntry` is created per advertised logical plugin, carrying the `factoryPluginId` that enumeration reported.

`EnsureLoaded()` performs the actual load and is idempotent (it returns `S_OK` immediately when `module`, `fileSystem`, and `informations` are already populated). It:

1. `LoadLibraryExW(path, nullptr, LOAD_WITH_ALTERED_SEARCH_PATH)`.
2. Registers the module as a localization resource owner (unregistered on any later failure via `wil::scope_exit`).
3. Resolves `RedSalamanderCreate` with `GetProcAddress` — a missing export fails the entry with `ERROR_PROC_NOT_FOUND`.
4. Calls the factory: `createFactory(__uuidof(IFileSystem), &options, GetHostServices(), entry.factoryPluginId.c_str(), fileSystem.put_void())`. The `IHost*` passed in is the host-services singleton.
5. `QueryInterface`s the instance for `IInformations` and reads `GetMetaData`. The reported `id` must match `factoryPluginId` (case-insensitive) or the entry fails with `E_FAIL`; `shortId` is validated by `IsValidShortId`.
6. Applies persisted JSON configuration via `ApplyConfigurationFromSettings` -> `IInformations::SetConfiguration`.

On success the entry is marked `loadable`, the COM pointers (`fileSystem`, `informations`) and the `wil::unique_hmodule` are stored, and the resource-owner registration is kept.

### Active plugin

Exactly one plugin is "active" per `_activePluginId`. `SetActivePlugin()` clears any disabled flag, calls `EnsureLoaded()`, then mirrors the id into `settings.plugins.currentFileSystemPluginId` and `SessionState`. `DisablePlugin()` of the active plugin first switches to the first other loadable, enabled plugin (or fails with `ERROR_ACCESS_DENIED` when none exists).

### Quiet point and unload

`Unload()` centralizes teardown in a strict order:

1. Release the COM instances first (`entry.informations = nullptr; entry.fileSystem = nullptr;`).
2. Call the optional `RedSalamanderPluginShutdown` export, if present. By contract this is the plugin's quiet point: after it returns, no DLL-global worker may call host callbacks or touch state that is about to be freed. It must be idempotent and non-throwing.
3. `Localization::UnregisterResourceOwner` for the module.
4. Decide how to release the `HMODULE` based on `ModuleUnloadMode`.

`ModuleUnloadMode` controls the final step:

- `FreeLibrary` — the normal path (runtime refresh, enable/disable, re-discovery). The module is released with `entry.module.reset()` so a fresh DLL image can be loaded later. The retain vote is **not** consulted here.
- `ProcessShutdown` — only here is `RedSalamanderPluginRetainModuleUntilProcessExit` queried. If it returns `TRUE`, the `HMODULE` is leaked (`entry.module.release()`) so process teardown does not race driver/global cleanup; otherwise it is released normally.

`Shutdown()` persists each plugin's configuration (`PersistConfigurationToSettings`, gated by `IInformations::SomethingToSave`), then calls `UnloadAll(ModuleUnloadMode::ProcessShutdown)`, which unloads entries in reverse order. Runtime `Refresh()` and the enable/disable/remove operations route through `EnsureLoaded`/`Unload` with `FreeLibrary` mode, which is why a refresh always releases the old image even if the plugin voted to retain at process exit.

## Host services and thread marshaling

`GetHostServices()` (`HostServices.cpp`) returns a never-deleted singleton (its `Release()` only decrements a refcount; it is process-lifetime). The class implements `IHost` plus the sibling interfaces from `Common/PlugInterfaces/Host.h`, each reachable via `QueryInterface`:

| Interface | Methods |
|---|---|
| `IHost` | Root object (extensible) |
| `IHostAlerts` | `ShowAlert`, `ClearAlert` |
| `IHostPrompts` | `ShowPrompt` |
| `IHostConnections` | Connection manager, secret get/set/prompt/clear, FTP upgrade |
| `IHostPaneExecute` | `ExecuteInActivePane` |
| `IHostViewers` | `OpenViewer` |

### The marshaling rule

Plugin code may call these methods on **any** worker thread, but the UI work has to run on the FolderWindow thread. Every method first resolves the host window (`GetInitializedHostWindow`) and checks `IsCurrentThreadWindowThread()`:

- **On the window thread** — call the real `*OnUiThread` handler directly.
- **Off the window thread** — build a heap payload struct (e.g. `PendingAlert`, `PendingPrompt`, `PendingConnectionSecret`) that owns deep copies of all caller strings, then hand it to the window.

Off-thread calls split by whether the caller needs a synchronous result:

| Direction | Transport | Examples |
|---|---|---|
| Fire-and-forget (no result needed) | `PostMessagePayload` (async; payload ownership transferred with `std::move`) | `ShowAlert`, `ClearAlert`, `ExecuteInActivePane` |
| Result-returning (must block) | `SendMessageW` (synchronous; window proc runs the handler and returns the `HRESULT` as the `LRESULT`) | `ShowPrompt`, `ShowConnectionManager`, `GetConnectionJsonUtf8`, `GetConnectionSecret`, `SetConnectionSecret`, `DeleteConnectionSecret`, `PromptForConnectionSecret`, `ClearCachedConnectionSecret`, `UpgradeFtpAnonymousToPassword`, `OpenViewer` |

The blocking `SendMessageW` calls pass the payload pointer in `lParam` and read back results (e.g. a `wil::unique_cotaskmem_string` secret) from the payload after the call returns. The async `PostMessagePayload` calls release ownership of a `std::unique_ptr` into the message queue.

### kHost* dispatch

The window proc calls `TryHandleHostServicesWindowMessage(message, wParam, lParam, result)`, which forwards to `HostServices::TryHandleMessage`. That method switches on the `WndMsg::kHost*` message ids and, for each, recovers the payload and invokes the matching `*OnUiThread` handler:

- Posted payloads are recovered with `TakeMessagePayload<T>(lParam)` (takes ownership and frees it).
- Sent payloads are recovered with a plain `reinterpret_cast<T*>(lParam)` (caller still owns the stack/heap object).

Message ids include `kHostShowAlert`, `kHostClearAlert`, `kHostShowPrompt`, `kHostShowConnectionManager`, `kHostGetConnectionJsonUtf8`, `kHostGetConnectionSecret`, `kHostSetConnectionSecret`, `kHostDeleteConnectionSecret`, `kHostPromptConnectionSecret`, `kHostClearCachedConnectionSecret`, `kHostUpgradeFtpAnonymousToPassword`, `kHostExecuteInPane`, and `kHostOpenViewer`.

### ABI validation

Every request struct carries `version` and `sizeBytes`. Handlers reject mismatches with `E_INVALIDARG` before reading any other field (`request->version != 1 || request->sizeBytes < sizeof(...)`). This is the same append-only ABI evolution contract documented in `Common/PlugInterfaces/Host.h` and `Specs/Plugins/Plugins_PluginAPI.md`.

## Capability gating for cross-FS copy/move

When a copy/move resolves a different `IFileSystem` instance for the source and destination panes, the host cannot delegate to a single plugin. The transfer is gated first, then executed by the bridge.

`IFileSystem::GetCapabilities` is **mandatory**: every provider must return `S_OK` and a parseable version-1 UTF-8 JSON document with `operations`, `concurrency`, `crossFileSystem`, and `pathIdentity`. The host-recognized shape (see `Common/PlugInterfaces/FileSystem.h` and `Specs/Plugins/Plugins_VirtualFileSystem.md`) is, abbreviated:

```json
{
  "version": 1,
  "operations": { "copy": true, "move": true, "read": true, "write": true, "delete": true },
  "concurrency": { "copyMoveMax": 4, "deleteMax": 8, "deleteRecycleBinMax": 2 },
  "crossFileSystem": {
    "export": { "copy": ["*"], "move": ["*"] },
    "import": { "copy": ["builtin/file-system"], "move": [] }
  },
  "pathIdentity": { }
}
```

`TryGetCapabilities` parses the document; any capability failure (failed `GetCapabilities`, null/empty JSON, unparseable JSON, or a document missing `operations`, `concurrency`, `crossFileSystem`, or `pathIdentity`) is **fail-closed**: it returns no capabilities, and `ReportCapabilitiesContractViolationOnce` logs the violation exactly once per provider instance via `Debug::Error`. Capability-gated operations are then disabled. Missing, invalid, unsupported, or unstable `pathIdentity` rejects identity-sensitive planners such as Batch Rename and same-context Copy/Move/Delete before mutation.

`CanCrossFileSystemCopyMove(sourceFs, sourcePluginId, destFs, destPluginId, operation)` (in `FolderWindow.FileOperations.cpp`) decides whether a cross-FS transfer is allowed:

1. The operation must be `FILESYSTEM_COPY` or `FILESYSTEM_MOVE`.
2. Both providers must return valid capabilities.
3. The source must advertise `read` and the destination must advertise `write`.
4. For `FILESYSTEM_MOVE`, the source must additionally advertise `delete` (the source copy has to be removed after a successful transfer).
5. The id allow-lists must agree in **both** directions: the source's `crossFileSystem.export.{copy,move}` list must allow the destination plugin id, and the destination's `crossFileSystem.import.{copy,move}` list must allow the source plugin id. `IdListAllows` treats `"*"` as a wildcard and matches ids case-insensitively.

Same-filesystem operations use the parallel `CanSameFileSystemOperation`. `FILESYSTEM_RENAME` is always allowed, but `FILESYSTEM_COPY`/`FILESYSTEM_MOVE`/`FILESYSTEM_DELETE` additionally require a **stable path identity** (`pathIdentity.pathTextStableIdentity: true`) in addition to the relevant `operations` flag.

> **Fail-closed `pathIdentity`.** `pathIdentity` is now a **mandatory** section of the capabilities document. `ParseFileSystemCapabilitiesJson` requires the block to be present and to pass `TryParseFileSystemPathIdentityContract`; if it is missing or fails the strict parser the **entire** capability set is rejected and **all** file operations for that provider (including same-FS copy/move/delete and cross-FS transfers) are refused. The strict parse requires: `version: 1`, `normalization: "none"`, a single-character `preferredSeparator` that is contained in `acceptedSeparators`, valid `componentComparison`/`caseOnlyRename` enum values, and the `casePreserving` + `pathTextStableIdentity` booleans. A plugin that ships a malformed `pathIdentity` is silently disabled — `PluginContractTests` asserts every shipped plugin's block parses under this exact contract so such a regression fails CI rather than reaching users.

## The cross-FS bridge

`CrossFileSystemBridge` (`FolderWindow.FileOperations.State.cpp`) performs the byte-level transfer. It `QueryInterface`s `IFileSystemIO` on both providers (`IFileSystemDirectoryOperations` on the destination for directory creation) and uses `IFileReader`/`IFileWriter` from `Common/PlugInterfaces/FileSystem.h`.

### Buffer sizing

The transfer buffer size is resolved by `ResolveAdaptiveCrossFsBridgeBufferBytes`, which combines:

- the configured `settings.fileOperations.crossFsBridgeBufferSizeKB` (default `4096` KB, clamped to `512`..`16384` KB), and
- per-endpoint `IFileSystem::GetTransferHints` (`preferredBufferBytes`) and `GetStorageCharacteristics`.

The resolved value is cached in `task._resolvedCrossFsBridgeBufferBytes` for diagnostics.

### Per-file transfer: temp + atomic commit

For each file, the bridge does **not** write directly to the final destination. Instead:

1. `MakeTempDestinationPath` builds a sibling temp name `<destination>.rs_tmp_<128-bit-hex>_<streamId>`. The 128 bits come from `BCryptGenRandom` (CSPRNG) so a local attacker cannot pre-create or race the staging file; the stream id is for diagnostic correlation only. (If `BCryptGenRandom` fails, it degrades to a PID/TID/tick name rather than blocking the copy.)
2. It opens an `IFileReader` on the source (capturing `GetSize` and `GetFileBasicInformation` first), and an `IFileWriter` on the temp path. The temp writer is opened with the overwrite/replace-readonly flags **stripped** (`tempFlags`), because the random name should never collide.
3. It pumps source bytes into the writer (serial or pipelined; see below).
4. After the last write, it checks the integrity invariant: when the source size was known, `fileCompletedBytes` must equal `fileTotalBytes`, otherwise the file fails with `ERROR_PARTIAL_COPY` (`bridge.integrity.sizeMismatch`).
5. `IFileWriter::Commit()` flushes the temp file; a failed commit is logged as `bridge.commit` and aborts the file.
6. `PromoteTempToFinalPath` renames the temp into the final destination via `destinationFs.MoveItem(temp, destination, promoteFlags, ...)`. The conflict resolution that was already answered for the destination (per-file or apply-to-all) is carried into `promoteFlags` as one-shot `FILESYSTEM_FLAG_ALLOW_OVERWRITE` / `FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY` grants.

A `wil::scope_exit` (`cleanupTemp`) calls `BestEffortDeleteTempFile` if the file was staged but not promoted, so a failed or cancelled transfer does not leave the staging file behind. `BestEffortDeleteTempFile` tolerates "file/path not found" and "invalid name" results silently.

After a successful promote:

- For `FILESYSTEM_MOVE` with a known source size, the bridge re-opens the destination, re-`GetSize`s it, and fails with `ERROR_PARTIAL_COPY` (preserving the source) if it cannot confirm the destination matches the source size — the move only deletes the source once the destination is verified.
- Source basic information (timestamps/attributes) is applied with `SetFileBasicInformation`; `E_NOTIMPL`/`ERROR_NOT_SUPPORTED` are tolerated.

### Serial vs producer/consumer pipeline

`ShouldUseBufferedPipeline(fileTotalBytes, bufferBytes)` decides the pump mode. In production it returns true only when the file is larger than one buffer (`bufferBytes > 0 && fileTotalBytes > bufferBytes`); under `ENABLE_TESTS`, a self-test override (`GetBridgePipelineModeOverride`) can force `Disabled` or `Enabled`.

- **Serial** (`copySerial`) — a single loop reads a buffer from the source then writes it to the destination, one buffer at a time. This is also the fallback when the secondary buffer cannot be allocated or the reader `std::jthread` cannot be created (`std::system_error`).
- **Producer/consumer** — two `BufferSlot`s and a second buffer enable double-buffering. A `std::jthread` reader fills the next slot (`pipelineReader->Read`) while the calling thread (the writer) drains the ready slot (`writer->Write`). Coordination uses one `std::mutex` + `std::condition_variable`; `pipelineStop`, `readerFinished`, and per-slot `ready` flags drive handoff. On cancel/error the writer sets `pipelineStop`, notifies, and `join()`s the reader before returning. Reader/writer wait and I/O times are accumulated into `BridgeCopyPerf` (`readerWaitUs`, `writerWaitUs`, `readUs`, `writeUs`).

Both modes share the same accounting: per-write they advance `fileCompletedBytes`, atomically advance the task's `overallCompletedBytes`, call `maybeReportProgress` (throttled to `ProgressIntervalMs` = 200 ms, forced when a file completes), and call `ThrottleThreadSafe` to honor any bandwidth limit. `WaitWhilePaused` and `CancelRequested` (which checks both `task._cancelled` and the `std::stop_token`) are polled in the inner loops so pause/cancel stay responsive.

### Directories, conflicts, and reparse points

Directory-vs-directory existence is a **merge**, not a conflict (per the `IFileSystem` copy/move contract in `FileSystem.h` and `Specs/FileSystem/FileSystem_FileOperations.md`): the bridge recurses into the existing destination directory and never fails the whole transfer on the first collision. File-vs-file (or type-mismatch) collisions are raised as **per-file** conflicts through `BridgeCallback` -> `IFileSystemCallback::FileSystemIssue`; the host serializes those prompts and caches apply-to-all answers. A file may not replace a directory through the bridge. Files answered Skip increment `skippedFileConflictCount` (the file is preserved at the source for a move) and the transfer ends `ERROR_PARTIAL_COPY`. Reparse points are handled per `reparsePointPolicy`.

### Progress, cancellation, and callbacks

The bridge reports progress through a `BridgeCallback` (an `IFileSystemCallback` adapter) guarded by `callbackMutex`. Each worker uses a distinct, stable `progressStreamId` so concurrent streams stay correlated — matching the `IFileSystemCallback` contract that parallel workers must report distinct stream ids and never invoke callbacks concurrently for one operation. Performance counters (`bridgeCopyUs`, `bridgeReadUs`, `bridgeWriteUs`, `bridgeReaderWaitUs`, `bridgeWriterWaitUs`, admission/queue depth counts) are emitted via `Debug::Perf::Emit` at completion.

## External file actions (view/edit launch)

When an item is opened with an external program rather than a built-in viewer, two helpers cooperate:

- `FileActionResolver` (`FileActionResolver.h`) resolves which `FileActionDefinition` applies. `ResolveViewerAction`/`ResolveEditorAction` take a `Request` (a `Command` such as `View`/`Edit`/`EditNew`, the item path, and the computer name) and return a `Resolution` whose `Reason` records why an action matched (e.g. `ComputerExtensionRule`, `GlobalDefaultRule`) or did not (`NoAssociation`, `ActionDisabled`, `ActionNotApplicable`). `CollectAssociatedEditorActions`/`CollectApplicableActions` and `ActionAppliesToContext` enumerate the per-context candidates.
- `FileActionLauncher` (`FileActionLauncher.h`) turns a resolved action into an actual process. `ExpandMacros` substitutes a `MacroContext` (item path, current/opposite-pane directories, a written selected-paths file, selected paths, computer name) into the template; `BuildExternalLaunchPlan` produces a `LaunchPlan` (executable, arguments, working directory, and any temp files to clean up after exit); `LaunchExternalPlan` runs it with `LaunchOptions` (owner window, show command, optional wait-for-exit and process-handle capture) and reports a `LaunchResult`.

## Gotchas

- **Capabilities are fail-closed.** A provider that returns `E_NOTIMPL`, an empty document, or an unparseable document from `GetCapabilities` gets every capability-gated operation disabled, and the violation is logged once per instance.
- **Cross-FS gating is bidirectional.** Both the source `export` allow-list and the destination `import` allow-list must agree; one side's `"*"` is not enough.
- **Never assume host-services UI is synchronous off-thread.** Prompts and secret reads block via `SendMessageW`; alerts and pane-execute do not block (`PostMessagePayload`). Calling an alert and immediately expecting the user to have responded is a bug.
- **The retain vote only applies at process shutdown.** Runtime refresh always releases the old module so a fresh DLL image can be loaded.
- **Atomic-commit only covers the final rename.** The bridge guarantees the destination path either has the fully written file or nothing (temp is cleaned up on failure); it does not make the whole multi-file operation transactional.

See also: [FileOperations.md](../FileOperations.md) for user-facing copy/move/queue behavior, [Plugins.md](../Plugins.md) and [RemoteFileSystems.md](../RemoteFileSystems.md) for the provider list, and `Specs/Plugins/Plugins_VirtualFileSystem.md` for the full capability and `pathIdentity` JSON contracts.
