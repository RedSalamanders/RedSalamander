# FileSystem P3 Architecture RFC (Callbacks + ABI Versioning + API Shape)

Last updated: 2026-02-27

## Scope

This RFC covers long-term interface and contract improvements for:

- `Common/PlugInterfaces/FileSystem.h` (filesystem plugin ABI + callbacks)
- host-side file operations (`RedSalamander/FolderWindow.FileOperations.*`)
- filesystem plugins (`Plugins/FileSystem*`)

It intentionally does **not** change behavior of existing operations; it defines a migration path.

## Problems

### 1) Callback lifetime is unsafe by construction

Current filesystem callbacks (`IFileSystemCallback`, `IFileSystemDirectoryWatchCallback`, `IFileSystemSearchCallback`) are **raw vtables** (not COM) and are described as host-owned pointers. This makes it easy to accidentally:

- store a callback pointer past the allowed lifetime,
- invoke callbacks after `UnwatchDirectory` / after an operation returns,
- race plugin background work against host teardown.

These are “design risks”: correctness depends on every plugin following the contract perfectly.

### 2) No explicit ABI versioning for most structs

Most structs crossing the host↔plugin boundary do not carry `version/sizeBytes` fields. This makes safe extension difficult (new fields require a lockstep rebuild of host + all plugins, with no runtime detection).

### 3) Single-item vs batch API duplication

The ABI exposes both single-item and batch operations:

- `CopyItem` vs `CopyItems`
- `MoveItem` vs `MoveItems`
- `DeleteItem` vs `DeleteItems`
- `RenameItem` vs `RenameItems`

This duplicates host and plugin logic and increases the chance of contract drift between the two paths.

## Goals

- Make callback lifetime safety **hard to violate**.
- Make struct evolution explicit via **ABI versioning**.
- Reduce duplicated entrypoints while preserving existing behavior and UX.
- Keep the migration path incremental: host can prefer v2 interfaces when available.

## Proposal

### A) Callback lifetime model

Introduce **v2 callback interfaces** that are COM (`IUnknown`) so plugins can safely take references when async work is required:

- `IFileSystemCallback2 : IUnknown`
- `IFileSystemDirectoryWatchCallback2 : IUnknown`
- `IFileSystemSearchCallback2 : IUnknown`

Rules:

- Host passes v2 callbacks when it supports them (via QI).
- Plugins may `AddRef()` and store v2 callbacks for async work.
- On unregister/teardown, the host can release its references; the plugin must stop producing and release its own references to reach a quiet point.

Migration:

- Keep existing raw-vtable callbacks for compatibility.
- Document a strict rule for v1 callbacks: **never store** the pointer beyond the call/registration scope.

### B) ABI versioning for structs

Adopt the `version/sizeBytes` pattern already used elsewhere (e.g., `HostPromptRequest` in `Specs/Plugins/Plugins_PluginAPI.md`):

- For each cross-boundary struct that may evolve, add:
  - `uint32_t version;`
  - `uint32_t sizeBytes;`
  - optional `reserved[]`

Migration options:

1) **ABI-breaking sweep**: update `Common/PlugInterfaces/FileSystem.h`, host, and all in-repo plugins together.
2) **Versioned additions** (preferred): add `*V2` structs and `*2` methods/interfaces so old binaries still load.

Recommendation: prefer (2) for long-lived external plugin compatibility; accept (1) only if the plugin ecosystem is strictly lockstep.

### C) Reduce entrypoint duplication

Define a new optional interface that collapses single + batch operations into one “batch-first” shape, while keeping today’s behavior:

- `IFileSystemBatchOps` (UUID new) with batch-only operations.
- Host implements single-item operations as thin wrappers that call into batch ops when available.
- Plugins may keep single-item methods for compatibility, but implement batch as the primary path.

Notes:

- `CopyItem/MoveItem` support explicit rename via `destinationPath`, while `CopyItems/MoveItems` use `destinationFolder + leaf`.
- A batch-first API should represent destination as either:
  - `(destinationFolder, leafOverride?)` per item, or
  - a per-item `(sourcePath, destinationPath)` pair.

## Acceptance / Next Steps

1) Decide migration strategy: ABI-breaking sweep vs versioned additions.
2) Define the minimal set of structs to version first (`FileSystemOptions`, watch/search progress structs, prompt structs).
3) Implement host-side preference order:
   - QI for v2 interfaces first, fall back to v1.
4) Add automated guards:
   - debug-only assertions that v1 callbacks are not used after unregister/return (generation counters / tokens).
5) Extend selftests to exercise both v1 and v2 paths once implemented.
