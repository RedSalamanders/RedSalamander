# FileSystem P3 Architecture RFC (Callbacks + ABI Versioning)

Last updated: 2026-03-01

## Scope

This RFC covers long-term interface and contract improvements for:

- `Common/PlugInterfaces/FileSystem.h` (filesystem plugin ABI + callbacks)
- host-side file operations (`RedSalamander/FolderWindow.FileOperations.*`)
- filesystem plugins (`Plugins/FileSystem*`)

It intentionally does not change behavior of existing operations; it formalizes contracts that existing implementations already follow and fills gaps where the contract is underspecified.

**Non-goal**: incremental v2 interface migration. All plugins are in the solution and will be adapted together.

## Problems

### 1) Callback lifetime is unsafe by construction

Current filesystem callbacks (`IFileSystemCallback`, `IFileSystemDirectoryWatchCallback`, `IFileSystemSearchCallback`, `IFileSystemDirectorySizeCallback`) are **raw vtables** (not COM) and are described as host-owned pointers. This makes it easy to accidentally:

- store a callback pointer past the allowed lifetime,
- invoke callbacks after `UnwatchDirectory` / after an operation returns,
- race plugin background work against host teardown.

These are "design risks": correctness depends on every plugin following the contract perfectly. The existing implementations (`FileSystem`, `FileSystemDummy`) handle this correctly via ad-hoc drain patterns, but the contract does not specify *how* — only that plugins "MUST NOT" invoke after return.

### 2) No explicit ABI versioning for most structs

Most structs crossing the host-plugin boundary do not carry `sizeBytes` fields. This makes safe extension difficult (new fields require a lockstep rebuild of host + all plugins, with no runtime detection).

`Host.h` structs (`HostAlertRequest`, `HostPromptRequest`, `HostConnectionManagerRequest`, etc.) already use explicit struct ABI versioning (`sizeBytes` + optional `reserved[]`). `FileSystem.h` structs do not.

## Proposal

### A) Callback drain protocol

Callbacks fall into two lifetime categories with different drain contracts.

#### Per-call callbacks

**Applies to**: `IFileSystemCallback`, `IFileSystemSearchCallback`, `IFileSystemDirectorySizeCallback`

The operation is blocking. When the call returns, the plugin has ceased all callback invocations — drain is implicit.

| Party | Obligation |
|-------|-----------|
| Plugin | MUST NOT invoke the callback after the operation returns. MUST check `ShouldCancel` at regular intervals and return promptly when cancelled. |
| Host | MUST keep the callback object and its backing state alive until the operation call returns. To tear down early, signal cancel via `ShouldCancel` and wait for the call to return before destroying state. |

No new mechanism needed. This is already how it works; the RFC formalizes it.

#### Registration-based callbacks

**Applies to**: `IFileSystemDirectoryWatchCallback`

`UnwatchDirectory` is the drain point. The "no callbacks after return" guarantee requires a synchronous drain:

| Party | Obligation |
|-------|-----------|
| Plugin | `UnwatchDirectory` MUST synchronously drain: (1) signal background work to stop, (2) wait for all in-flight callback invocations to **complete** before returning. |
| Host | MUST NOT destroy callback-referenced state until `UnwatchDirectory` returns. The callback implementation MUST be safe to invoke from any thread at any time before `UnwatchDirectory` returns. |

**Reference drain patterns** (both already implemented in-tree):

1. **Thread pool wait** (`FileSystem` plugin, `FileSystem.Watch.cpp`):
   - Set `_stopping` atomic flag
   - `CancelIoEx` on the directory handle
   - `WaitForThreadpoolIoCallbacks(tpIo, TRUE)` — blocks until all I/O callbacks complete
   - `WaitForThreadpoolWorkCallbacks(tpWork, TRUE)` — blocks until all work callbacks complete
   - All callback paths check `_stopping` and exit early

2. **Atomic counter + condition variable** (`FileSystemDummy` plugin):
   - Set `active` atomic flag to `false`
   - Wait on `_watchCv` until `inFlight` counter reaches 0
   - Callback entry increments `inFlight`; callback exit decrements and notifies CV

**Reentrant unwatch**: If a watch callback triggers host-side navigation that calls `UnwatchDirectory` on the same registration (reentrant), the drain must account for the active callback on the current call stack. `FileSystemDummy` handles this via `g_activeDirectoryWatchCallback` — the drain waits for `inFlight <= 1` instead of `inFlight == 0` when reentrant. This is an allowed pattern.

### B) Thread safety requirements

#### Callback invocation threading (existing, formalized)

- All callbacks may be invoked from any thread (background workers, thread pool, I/O completion threads).
- Per-call: plugin MUST serialize invocations per-operation (one callback in flight at a time).
- Watch: plugin MUST serialize invocations per-registration.

#### Host callback implementation requirements (new)

- Host callback implementations MUST tolerate invocation racing with a teardown request. A callback arriving between "host decides to tear down" and "drain completes" must not cause UB.
- The standard pattern is an atomic `_stopping` flag checked at callback entry, as `FolderWatcher::OnPluginDirectoryChanged` already does.
- Callback implementations MUST NOT assume they run on any particular thread.
- Callback implementations MUST NOT take locks that the teardown path holds (or vice versa), to avoid deadlock during the drain wait.

#### Plugin drain implementation requirements (new)

- The drain wait in `UnwatchDirectory` MUST NOT hold the same lock that callback delivery acquires. Both existing implementations handle this correctly: `FileSystem` unlocks `_mutex` before calling `WaitForThreadpool*Callbacks`; `FileSystemDummy` uses a CV wait that releases the lock.
- The `_stopping` / `active` flag MUST be set **before** initiating the drain wait, so in-flight callbacks can see it and exit promptly.

#### Deadlock avoidance rule

Watch callbacks MUST NOT perform synchronous calls that depend on the thread that will call `UnwatchDirectory`. In practice:

- Use `PostMessage` / `TrySubmitThreadpoolCallback`, never `SendMessage`, from a watch callback.
- `FolderWatcher::OnPluginDirectoryChanged` already does this correctly (posts to thread pool).

If `UnwatchDirectory` is called from the UI thread and a watch callback does `SendMessage` to the UI thread, classic deadlock results:
- UI thread blocks in `UnwatchDirectory` waiting for drain
- Plugin background thread is in callback, blocked on `SendMessage` to UI thread

### C) ABI versioning for structs

Adopt a `sizeBytes` field at the start of structs that cross the host↔plugin boundary.

- For each cross-boundary struct that may evolve, add:
  - `uint32_t sizeBytes; // sizeof(StructName)`
  - optional `reserved[]` (only when we want extra pre-allocated space for future fields without changing the struct size)

**Rules (normative):**

- **Initialization**:
  - The **creator** of the struct MUST set `sizeBytes = sizeof(StructName)` before passing it across the boundary.
  - Any optional `reserved[]` fields MUST be zero-initialized.
- **Consumption**:
  - The **consumer** MUST validate `sizeBytes` before reading fields.
  - For this RFC’s **ABI-breaking sweep**, `sizeBytes != sizeof(StructName)` is treated as a contract violation (debug-assert and fail the operation with `E_INVALIDARG` where applicable).
- **Out parameters**:
  - For `[out]` structs (e.g. `FileSystemDirectorySizeResult* result`), the **caller** MUST initialize `sizeBytes` before calling into the plugin.
  - The plugin MUST treat `sizeBytes` as the writable size and MUST NOT write beyond it.

Migration: **ABI-breaking sweep**. Update `Common/PlugInterfaces/FileSystem.h`, host, and all in-repo plugins together. All plugins are in the solution; no external compatibility needed.

**Priority order for structs to version**:

| Struct | Urgency | Rationale |
|--------|---------|-----------|
| `FileSystemOptions` | High | Only has one field (`bandwidthLimitBytesPerSecond`), most likely to grow. Passed as in/out through every callback. |
| `FileSystemSearchQuery` | High | Likely to gain content-search fields, encoding options, etc. |
| `FileSystemSearchMatch` | Medium | May gain additional metadata fields. |
| `FileSystemSearchProgress` | Medium | May gain additional counters. |
| `FileSystemDirectoryChangeNotification` | Medium | May gain filter/scope fields. |
| `FileSystemDirectoryChange` | Deferred | Array element inside `FileSystemDirectoryChangeNotification`. If it ever needs to grow, we’ll need per-entry `sizeBytes` or a stride field on the notification. |
| `FileSystemDirectorySizeResult` | Low | Stable shape, unlikely to change soon. |
| `FileSystemBasicInformation` | Low | Crosses the boundary via `IFileSystemIO::GetFileBasicInformation` / `SetFileBasicInformation`. Stable shape, but should be listed for completeness. |
| `FileSystemRenamePair` | Low | Simple two-pointer struct, unlikely to change. |
| `FileInfo` | Deferred | Uses linked-list `NextEntryOffset` layout; versioning requires a different approach (e.g. `sizeBytes` on `IFilesInformation`, not per-entry). |

### D) Debug-only validation

Add host-side active-scope guards / dead flags to catch contract violations in debug builds. Zero-cost in release builds.

#### Per-call callbacks

**Host-side guard (no plugin cooperation required):**

- Host callback implementations keep a debug-only “active scope” counter.
- The host increments the counter before calling into a synchronous plugin operation that may invoke callbacks, and decrements it after the operation returns.
- Every callback entry debug-asserts that the active-scope counter is non-zero.

This catches:
- callbacks invoked after a synchronous operation returned (use-after-return contract violation),
- callbacks racing with teardown paths where host state is no longer valid.

#### Registration-based callbacks

- Plugin embeds a `std::atomic<bool> dead` flag in the watch registration.
- `UnwatchDirectory` sets `dead = true` after the drain completes.
- Debug-assert `!dead` at callback entry.

## Acceptance

- [x] ABI-breaking sweep for struct versioning — add `sizeBytes` to cross-boundary structs per priority table above (add `reserved[]` only when needed).
- [x] Document the callback drain protocol in `FileSystem.h` header comments (Proposal A contracts).
- [x] Document thread safety requirements in `FileSystem.h` header comments (Proposal B contracts).
- [x] Add debug-only host-side active-scope guards for per-call callbacks and dead flags for watch callbacks (Proposal D).
- [x] Extend selftests to verify drain behavior: test that callbacks are not invoked after operation return / `UnwatchDirectory` return.
