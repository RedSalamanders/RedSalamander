# MTP/PTP File System

## Scope

`FileSystemMtp.dll` exposes Windows Portable Devices (MTP/PTP phones, cameras, and media players) as a RedSalamander virtual file system through the in-box Windows Portable Devices (WPD) COM API.

Canonical plugin identity:
- Long ID: `builtin/file-system-mtp`
- Short ID: `mtp`
- Embedded protocol token: `IDS_FILESYSTEMMTP_FSNAME` (`MTP`), intentionally language-neutral

## Identity And Paths

The plugin accepts `mtp:/`, `mtp://`, slash-rooted paths, and connection-manager paths, then normalizes them to one canonical slash-rooted provider path.

Connection Manager profiles use:
- `pluginId`: `builtin/file-system-mtp`.
- `host`: required WPD PnP device id reconnect hint.
- `initialPath`: optional path below the selected device root; default `/`.
- `authMode`: `anonymous`; MTP has no user/password/secret.
- `extra.readOnly`: defaults to `true`. Setting it to `false` opts the profile into production WPD mutations when the user has approved the target device/storage.
- `extra.devicePuid` and `extra.friendlyName`: optional picker-provided stable/display hints.

The Connection Manager MTP editor is a no-auth surface. It must hide username,
password/secret, raw host, and raw initial-path fields, and must show a
worker-backed device picker, storage picker, and `readOnly` toggle instead. The
UI thread must not synchronously enumerate WPD devices. Picker workers must
route device and storage browse through the MTP plugin factory contract
(`RedSalamanderBrowseConnectionTargets`) keyed by `pluginId`; the host must not
carry its own WPD picker stack or host-side picker fixtures. The MTP plugin owns
both live WPD browse and deterministic fake-backend browse for selftests, returns
plain picker JSON allocated with `CoTaskMemAlloc`, and emits
`mtp.connection_browse.devices*` / `mtp.connection_browse.storages*` metrics for
the browse calls. The Connection Manager worker then marshals plain picker data
back with `PostMessagePayload`, consumes it with `TakeMessagePayload`, and drains
queued picker payloads when the window is destroyed. Saving an MTP profile stores
the selected WPD PnP id in `host`, keeps `authMode` anonymous, stores picker
display identity in `extra.friendlyName`, stores the best available device
identity in `extra.devicePuid`, and defaults `extra.readOnly` to `true`. The
plugin-backed picker must prefer the device object's
`WPD_OBJECT_PERSISTENT_UNIQUE_ID` for `extra.devicePuid` when WPD exposes it, and
may fall back to the PnP id only when the persistent id is unavailable.

Device root components must include a stable identity suffix:
- `devpuid:<id>` when WPD exposes a persistent device identifier.
- `devid:<id>` when only a session-scoped device identifier is available.

Object path components must be stable even when sibling display names duplicate:
- Use `puid:<id>` when the object exposes a persistent unique id.
- Use `oid:<id>` only as a session-scoped fallback.

MTP identity formatting must be centralized in the plugin shared helpers. Device
root suffixes, persistent/session object suffixes, duplicate-name suffixes,
connection-root hashes, overwrite-journal device-state directory hashes, path
component sanitization, and plugin JSON escaping must flow through
`FileSystemMtp.Shared.cpp` helpers declared in `FileSystemMtp.Internal.h`
(`StableMtpIdentityHash`, `FormatMtpIdentityHash`,
`SanitizeMtpPathComponent`, `Mtp*IdentitySuffix`,
`MtpDuplicateObjectSuffix`, and `JsonEscapeUtf8`). Core, live WPD backend,
fake backend, and factory browse JSON code must not carry local hash,
sanitizer, suffix-format, or JSON-escape clones.

Unsuffixed paths may be accepted for browsing only when they resolve unambiguously. Any mutation against an ambiguous unsuffixed component must fail closed instead of selecting an arbitrary object.

MTP path identity is `ordinalCaseSensitive`. The provider path text is the mutation/cache identity; localized labels are display-only.

Case-insensitive MTP identity comparisons that intentionally match device/session
state rather than provider path text must use the shared `OrdinalString` helpers.
Device/session/cache-key equality uses `OrdinalString::EqualsNoCase`; stable
case-folded cache keys use `OrdinalString::FoldCaseInvariant`. Do not add local
CRT `towlower`/`towupper` equality loops for these comparisons. This does not
change provider path identity: exposed MTP paths remain `ordinalCaseSensitive`.

Duplicate sibling groups must be instrumented and testable without a live device. The deterministic `mtp_duplicate_names_require_stable_suffix` selftest guards the fake-backend path by checking that duplicate siblings are exposed only through stable `puid:` suffixes, suffixed duplicate paths read distinct contents, literal names containing suffix-like text plus `%`, `]`, trailing-space, and non-ASCII names survive enumeration, repeated enumeration keeps the same suffixed entries, and an unsuffixed ambiguous mutation fails closed.

## Read-Only WPD Traversal

The live WPD backend enumerates root devices with `IPortableDeviceManager::GetDevices`. Device root names include a `devid:` suffix derived from the WPD PnP device id when no persistent device id is available.

Below a device root, paths are resolved by walking WPD object ids from `WPD_DEVICE_OBJECT_ID` through `IPortableDeviceContent::EnumObjects`. Display names prefer `WPD_OBJECT_ORIGINAL_FILE_NAME`, then `WPD_OBJECT_NAME`, then the object id. Sibling display names that collide case-insensitively must be made distinct before exposure by appending a `puid:` suffix when `WPD_OBJECT_PERSISTENT_UNIQUE_ID` exists, otherwise an `oid:` suffix.

The live backend opens devices with `CLSID_PortableDeviceFTM`. Read-only calls request `WPD_CLIENT_DESIRED_ACCESS = GENERIC_READ`; mutating calls request `GENERIC_READ | GENERIC_WRITE`. It reads file-like objects through `IPortableDeviceResources::GetStream(..., WPD_RESOURCE_DEFAULT, STGM_READ, ...)`. When the profile remains `readOnly:true`, the core entry point blocks writer/create/delete/rename/copy/move attempts before WPD is touched. When the profile explicitly sets `readOnly:false`, production WPD supports whole-object file upload, folder creation, delete, same-parent rename, native same-device same-leaf copy/move where the device accepts it, and file copy/move fallback by read-stream plus upload. Direct backend overwrite remains rejected; safe overwrite is owned by the core temp/PUID/journal flow.

When initialized from a `host` PnP hint, the plugin resolves connection-root calls through a synthetic `devid:` root component derived from the same stable hash used by the live WPD enumerator. The WPD resolver must match `devid:` tokens independently of the friendly-name prefix so a device rename does not make a saved profile unusable.

## Reader Seek Contract

`CreateFileReader` is streaming-first. Opening a reader must resolve the object and prepare a backend stream without downloading the whole file; `GetSize` returns the reported object size without reading file contents. The deterministic `mtp_reader_streams_on_read_not_open` selftest guards this contract by requiring fake-backend read accounting to remain zero after `CreateFileReader` and `GetSize`, and to increment only when `Read` consumes bytes.

For seekable WPD streams, the returned `IFileReader` forwards `Read`, `Seek`, and `GetSize` through the serialized backend command queue while keeping the WPD stream/content/session alive. `Seek` supports `FILE_BEGIN`, `FILE_CURRENT`, and `FILE_END`, including backward seeks and positions beyond EOF. Reads beyond EOF succeed with `bytesRead == 0`. Null output pointers return `E_POINTER`; invalid seek origins return `E_INVALIDARG`; negative seeks before byte 0 return `ERROR_NEGATIVE_SEEK`.

If a device exposes a readable object but the WPD stream refuses seeking at open, the backend may fall back to a memory-backed reader to preserve the public `IFileReader` seek contract. This fallback is exceptional; ordinary readable, seekable MTP objects must not gate the first byte on a full-device download.

Public MTP helper COM objects that call back into their owning `FileSystemMtp`
instance, including file readers and staged writers, must keep that owner alive
with `wil::com_ptr<FileSystemMtp>` rather than owning raw pointers with manual
`AddRef`/`Release`. Directory enumeration result construction must keep
`FilesInformationMtp` in smart-pointer ownership until the successful
`IFilesInformation**` handoff, so failure paths do not manually release the COM
object.

## Capabilities

MTP is a high-latency, virtual, device-backed provider. Capabilities must be honest for the current device/profile:
- No random writes.
- No alternate data streams.
- No native v1 directory watch interface.
- Same-device rename/copy/move/delete/create-directory support is device-dependent and must fail with a specific HRESULT when unsupported or rejected.
- Public write APIs must stage data until `Commit`; the plugin must not create a final WPD object before commit.

After the user invokes the MTP Disconnect drive-menu command, the instance must stop using the old device session. Future device-backed operations must fail with `ERROR_DEVICE_NOT_CONNECTED`, and capability JSON must no longer advertise `operations.read`, `operations.properties`, or `operations.write` as available for that disconnected instance.

## Drive Info And Disconnect Menu

`IDriveInfo::GetDriveInfo` must expose the device-root segment as both the display name and volume label, and must expose the language-neutral `MTP` protocol identity as `fileSystem`. MTP v1 must not synthesize total/free/used capacity values when the storage capacity is not known.

`GetDriveMenuItems` must return a disabled Properties command and an enabled Disconnect command. Properties is unsupported in v1 and returns `ERROR_NOT_SUPPORTED`. Disconnect is not OS safe-removal; it is a plugin session release that is allowed even on read-only profiles. The deterministic `mtp_drive_info_and_disconnect_menu` selftest guards the fake-backend path by checking the drive labels, menu command ids/flags, unsupported Properties HRESULT, successful Disconnect, device-gone enumeration after Disconnect, and post-Disconnect capability booleans.

## Copy And Move Fallbacks

Native WPD copy/move support is optional. When the plugin falls back to download/upload for move, the delete-source phase is the only destructive part. If the destination upload/copy succeeds but deleting the source fails, `MoveItem` must:
- Return a specific failure HRESULT from the delete-source phase.
- Report the same failure through the item-completion callback when one is supplied.
- Leave both source and destination readable.
- Never claim success and never destructively roll back the destination.
- Emit transfer accounting for the copied bytes.

The deterministic `mtp_move_fallback_delete_source_failure_leaves_duplicate_and_reports_partial` selftest guards the fake-backend scaffold for this contract. Production WPD fallback copy/move is implemented for files by reading the source stream and uploading a destination object. Directory copy/move fallback remains native-only in v1; renamed directory copy/move that cannot be represented as a native same-leaf WPD operation returns `ERROR_NOT_SUPPORTED`. Mid-stream WPD cancellation remains part of closeout.

Successful copy operations must report initial progress, complete the item exactly once, leave a readable destination, and emit copied-byte accounting. The deterministic `mtp_copy_to_device_accounting` selftest guards the fake-backend path by checking the completion callback, copied contents, backend copy-call count, and `mtp.transfer.copy_bytes`.

Device-to-host reads through `CreateFileReader` must emit deterministic byte accounting per actual `Read` transfer, not during reader open or `GetSize`. The deterministic `mtp_copy_from_device_accounting` selftest guards the fake-backend path by checking the returned fixture payload, exactly one backend read, `lastReadBytes`, and `mtp.transfer.read_bytes`; `mtp_reader_streams_on_read_not_open` guards that opening and sizing the reader do not materialize bytes.

Copy/move transfer entry must remain cancellation-friendly. Before entering the serialized backend/device transfer, the plugin must emit an initial progress callback when one is supplied, poll `FileSystemShouldCancel`, and return `ERROR_CANCELLED` without creating the destination when the host has already cancelled. The deterministic `mtp_transfer_cancel_is_prompt` selftest guards this pre-transfer path with a delayed fake backend. This is not a substitute for the later WPD mid-stream `Revert`/watchdog path required for production writes.

Batch copy/move/delete/rename wrappers must report the host's per-item index in
each completion callback. They must not collapse delegated single-item
operations to item index `0`, because host batch engines key per-item results by
that index. The deterministic `mtp_batch_callbacks_report_item_indices`
selftest guards the fake-backend path with a two-item batch and requires
separate successful completions for item indices `0` and `1`.

## Transport Serialization

MTP/WPD is treated as a single-transport provider. `GetCapabilities().concurrency` must advertise `copyMoveMax:1`, `deleteMax:1`, and `deleteRecycleBinMax:0`. The plugin enforces this locally by executing every backend/device operation for one `FileSystemMtp` instance on its single command-queue worker; no second transport mutex is required around that already-serialized worker. Progress callbacks, cancellation callbacks, JSON buffer storage, and host notifications remain outside backend execution.

Each `FileSystemMtp` instance must submit backend/device work through one long-lived command queue worker rather than creating a fresh thread per filesystem call. The worker initializes COM as MTA once for its lifetime, serializes commands, replays any retained overwrite journal before the queued command, and reports completion back to the caller for watchdog waiting. Worker construction failure is surfaced through filesystem commands rather than terminating a `noexcept` constructor. On a watchdog timeout or explicit disconnect cleanup, the active worker is detached into the module unload quarantine, pending queued commands are completed with `ERROR_DEVICE_NOT_CONNECTED`, backend cancellation is requested, and unload remains deferred until the quarantined worker exits. The deterministic `mtp_backend_command_worker_is_reused` selftest guards that concurrent fake-backend metadata calls reuse one backend worker thread while preserving `maxConcurrentBackendCalls == 1`.

The live WPD backend must reuse device sessions across serialized calls instead of reopening WPD for every path. It caches `IPortableDevice`/`IPortableDeviceContent` by PnP id, upgrades the cached session when a later mutating command needs write access, and memoizes normalized path to object-id resolution for the active session. Successful enumeration replaces cached descendants with the returned child metadata. Size-sensitive transfers discard the target entry and resolve fresh metadata; a stream size mismatch invalidates and retries once. A missing-object result from a cache-resolved id evicts that subtree and retries once, while unexpected operation failures evict the affected PnP session and path entries. Successful writes invalidate the affected subtree; watchdog abandonment and Disconnect make the owning instance stop using the old session. The deterministic `mtp_wpd_session_and_path_cache_reuse` and `mtp_wpd_cache_failure_reopens_session_and_refreshes_size` selftests guard reuse, session-death recovery, fixture-option validation, and changed-size freshness.

## Directory Size And Progress

`GetDirectorySize` is high-metadata-cost on MTP and must remain cancellation-friendly. When a callback is supplied, the plugin must:
- Emit progress with current totals and the current path before checking `DirectorySizeShouldCancel`.
- Pass the host `cookie` back verbatim and never retain the callback after the call returns.
- Return the first progress/cancel callback failure as both the method HRESULT and `result.status`.
- On successful completion, emit one final progress callback with `currentPath == nullptr` and final totals before returning.

For a single-file `GetDirectorySize` target, the plugin must use metadata size
(`GetFileSize` / reported object size) rather than materializing the whole file
through `ReadFile`. The deterministic fake-backend coverage in
`mtp_fake_backend_enumerate_read_and_capabilities` checks that the single-file
directory-size path increments file-size accounting without increasing
fake-backend read-file calls.

## JSON Return Storage

`GetConfiguration`, `GetCapabilities`, and `GetItemProperties` must use per-instance, per-method two-slot UTF-8 return buffers. Each call writes into the inactive slot and returns that slot's `c_str()`. The implementation must reuse slot storage rather than retaining a per-item list, LRU ring, or unbounded cache. Returned pointers remain valid only until the next call to the same method or object release.

## Writes And Overwrite Safety

WPD transfers are whole-object operations. The safe overwrite model is:
- Stage new bytes under plugin control.
- Releasing a public writer without `Commit` is an abort and must leave no backend object behind.
- Record overwrite intent before commit when replacing an existing object.
- Require a non-empty temporary object identity before deleting or replacing the original.
- Default verification level is `transmitHash`: verify bytes sent by the plugin plus post-commit size. This is local transmit integrity only.
- `deviceReread` is the only device-storage verification guarantee and must be explicit opt-in.
- `sizeOnly` is allowed only when the profile accepts weaker verification.

The host and plugin must never present `transmitHash` as proof that the device stored identical bytes.

The current scaffold enforces the configured verification-accounting path during overwrite commits: every overwrite checks the committed object size, `transmitHash` additionally hashes the staged bytes without a device re-read, `deviceReread` performs one full re-read and compares bytes, and `sizeOnly` performs no byte hash or re-read. Writer overwrites that replace an existing file now use a GUID-named temporary sibling, verify that temp object, require a non-empty temp PUID from item properties before deleting the original, delete the original, and rename the temp into the final path on the success path. The deterministic `mtp_writer_overwrite_uses_temp_puid_swap` selftest guards replacement contents, final-path PUID handoff, exactly one final sibling, no leaked temp sibling, and the `mtp.overwrite.temp_puid_present` / `mtp.overwrite.temp_swap_committed` metrics.

If the writer overwrite temp upload fails after the journal intent was recorded but before the original is deleted, the overwrite must fail without changing the original, leak no temp sibling, clear the journal, and allow a later retry to proceed normally. The deterministic `mtp_overwrite_temp_upload_failure_keeps_original_and_allows_retry` selftest guards the fake-backend injected temp-upload failure branch and the `mtp.overwrite.temp_upload_failed` metric.

If a committed temp exposes no persistent ID, the writer overwrite must fail before deleting the original. The plugin must delete the harmless random temp when possible, leave the original bytes and final entry intact, record created-object PUID support as unsupported for that instance/device, and reject later existing-file overwrite attempts before uploading another temp. The deterministic `mtp_overwrite_empty_temp_puid_keeps_original_and_blocks_later_upload` selftest guards the fake-backend PUID-less branch and the `mtp.overwrite.temp_puid_missing` / `mtp.overwrite.temp_puid_policy_blocked` metrics.

If original deletion fails after the temp was verified and identified, the writer overwrite must fail without losing the original and must remove the random temp. A later retry may proceed normally. The deterministic `mtp_overwrite_delete_original_failure_keeps_original_and_allows_retry` selftest guards original-byte preservation, temp cleanup, single final entry visibility, retry success, and `mtp.overwrite.delete_original_failed`.

Device-sourced copy/move overwrites that replace an existing destination must use the same safe temp-sibling swap model rather than direct backend overwrite: copy the source to a GUID temp sibling without overwrite, verify the temp using the configured device-source policy, require a non-empty temp PUID, delete the original destination, rename the temp into the final destination path, and delete the source only after the final-path swap when the operation is a move. The deterministic `mtp_copy_move_overwrite_uses_temp_puid_swap` selftest guards copy and move replacement contents, final-path PUID handoff, copy-source retention, move-source deletion, exactly one final sibling, no leaked temp sibling, and `mtp.overwrite.device_source_temp_swap_committed`.

The temp-sibling overwrite protocol is file-only. Before recording journal intent or mutating the device, copy, move, and rename with `FILESYSTEM_FLAG_ALLOW_OVERWRITE` must read both source and existing-destination attributes and fail with `ERROR_ACCESS_DENIED` when either object is a directory. The deterministic `mtp_fake_backend_move_rejects_directory_transfer_fallback` selftest also exercises copy, move, and rename of a directory over an existing file and guards that the source tree and destination bytes remain unchanged.

Public `RenameItem` with `FILESYSTEM_FLAG_ALLOW_OVERWRITE` onto an existing destination is also a device-sourced overwrite and must use the temp-sibling swap path instead of forwarding backend `allowOverwrite`. Missing-destination renames call the backend with overwrite disabled, matching WPD's destination-must-not-exist contract. The deterministic `mtp_rename_overwrite_uses_temp_puid_swap` selftest guards replacement contents, source deletion, final-path PUID handoff from the temp object rather than the source object, no leaked temp sibling, and fake-backend `copyItemCalls` evidence that the overwrite used the temp-copy path.

Backend `MoveItem` implementations must match WPD's direct-device contract: destination must not already exist, directory moves have no transfer fallback when the native leaf-preserving move path is unavailable, and non-leaf-preserving directory moves return `ERROR_NOT_SUPPORTED`. File moves that cannot use the native leaf-preserving path may fall back to transfer-copy plus source delete. The fake backend mirrors this contract so fake-backed selftests cannot certify directory move behavior that WPD hardware rejects. The deterministic `mtp_fake_backend_move_rejects_directory_transfer_fallback` selftest guards the non-leaf-preserving directory case.

If device-sourced copy/move overwrite temp creation fails after the journal intent was recorded but before destination deletion, the overwrite must fail without changing source or destination, leak no temp sibling, clear the journal, and allow a later retry to proceed normally. For a retry, copy must retain the source and move must delete the source only after the swap succeeds. The deterministic `mtp_copy_move_overwrite_temp_copy_failure_keeps_original_and_allows_retry` selftest guards this branch for both copy and move and the `mtp.overwrite.temp_copy_failed` metric.

Host-side overwrite journal intent must be recorded before touching the device with an overwrite temp upload/copy for writer, copy, and move overwrite paths. If that step-0 journal write fails, the overwrite must fail closed before any backend write/copy/move, leave the source and destination contents unchanged, and leak no random temp sibling. The deterministic `mtp_overwrite_journal_write_failure_aborts_before_upload` selftest guards this preflight path for writer, copy, and move overwrites.

Before the next serialized backend command for a device, a retained overwrite journal must be replayed when the host-side journal path is available. The replay path handles the writer rename-temp failure window: destination missing with temp present renames the temp to the final path and clears the journal; destination present with temp present deletes the temp and clears the journal; destination present with no temp and no orphan-sweep requirement clears a stale terminal journal; both missing is recorded as a bounded terminal state and clears the journal. The deterministic `mtp_overwrite_journal_recovers_rename_temp_failure` selftest guards the temp-to-final replay branch after a one-shot rename failure, and `mtp_overwrite_journal_replay_removes_temp_when_final_exists` guards the final-present/temp-present cleanup branch.

If destination-present/temp-present cleanup cannot delete the random temp on a replay attempt, the plugin must retain the journal so the next serialized backend command retries cleanup instead of silently abandoning the temp. The deterministic `mtp_overwrite_journal_replay_temp_cleanup_delete_failure_retries` selftest guards the one-shot cleanup delete failure branch by requiring the first replay to fail and keep the journal, the next replay to remove the temp, preserve final bytes, and clear the journal.

If the destination exists, the exact temp path is missing, and a retained journal lacks `tempPuid` but has phase `planned`, `uploaded`, or `committed` plus `declaredSizeBytes` and `journalTimestampFileTimeUtc`, replay must not blind-delete by name. It sweeps the parent for streamable plugin temp objects matching the random temp leaf token, declared size, and timestamp. Exactly one candidate is deleted and clears the journal. If zero candidates remain and the destination size matches `declaredSizeBytes`, replay infers that the temp swap already completed before the journal clear, emits `mtp.overwrite.journal_replay_completed_swap_inferred`, and clears the journal. Zero candidates with a size mismatch, or multiple candidates, delete nothing and retain the journal for bounded retry. The deterministic `mtp_overwrite_journal_recovers_committed_temp_without_tempPuid` selftest guards the exact-one and ambiguous paths, and `mtp_overwrite_journal_clears_completed_swap_without_temp` guards the completed-swap inference path.

Retained no-tempPUID journals must carry a persisted `replayAttemptCount`. Replays that still cannot identify a safe terminal action rewrite the journal with an incremented count and emit `mtp.overwrite.journal_replay_stale_retained`; when the retained replay limit is reached, the plugin renames the journal to `overwrite-journal.json.stale`, emits `mtp.overwrite.journal_replay_stale_quarantined`, marks the identity as journal-absent in memory, and proceeds without replaying that stale journal on every later command. A quarantine rename failure emits `mtp.overwrite.journal_replay_stale_quarantine_failed` and leaves the journal for later retry.

If the recovery rename itself is repeatedly rejected, replay must not infinite-retry forever across opens. The current scaffold records a replay attempt count in the journal, schedules bounded retries for the first rejected recovery renames, and on retry exhaustion clears the journal while leaving the verified data available under the temp name. The deterministic `mtp_overwrite_journal_replay_rename_rejection_is_bounded` selftest guards the retry, terminal-clear, post-terminal no-retry, and retained-temp-readback behavior. User-facing recovery notification for this terminal temp-name state remains part of the broader closeout work if required by the product UX.

The deterministic `mtp_overwrite_never_duplicates_or_halfwrites` selftest aggregates the fake-backend overwrite crash-window matrix across temp upload failure, step-0 journal write failure, empty tempPUID policy block, rename-temp recovery, after-Commit orphan cleanup, and bounded unrenamable-temp terminal retention. It guards the no-duplicate-final and no-half-written-replacement contract for the deterministic scaffold; live WPD/device-specific write gates, user-facing terminal recovery UX if needed, and full package/validation closeout remain required before v1 closeout.

Device-sourced copy/move overwrites must not reuse the staged-writer `transmitHash` path over a just-read device temp because that would be circular. For those overwrites the default `transmitHash` path explicitly trusts the device commit and emits `mtp.verify.device_source_trust_commit`; `deviceReread` snapshots the device source before mutation, re-reads the destination after mutation, compares bytes, and emits `mtp.verify.device_source_bytes` plus `mtp.verify.device_source_reread_bytes`. The deterministic `mtp_overwrite_verify_input_by_source_kind` selftest guards both copy and move for the default and opt-in paths.

All command kinds may cache an absent overwrite journal per normalized, case-folded journal-storage root and device identity after a missing local journal probe. Including the storage root prevents an absent observation from one redirected/portable host session from suppressing replay after the journal root changes. Each identity owns a generation counter: intent recording increments the generation before invalidation/write, and an absent observation is accepted only if its observed generation is still current. This makes absent-cache reuse safe for mutating commands without hiding a journal created by another instance. Successful clear/quarantine marks the current generation absent. The deterministic `mtp_overwrite_journal_generation_and_absent_cache_are_constant_cost` selftest rejects a stale absent mark and proves sixteen deletes perform at most one filesystem journal probe; the full ordered replay matrix proves journals remain visible after the storage root changes between cases.

## Lifetime And Unload

MTP/WPD calls can wedge in drivers. Runtime refresh must use the common plugin lifetime contract:
1. Release normal COM instances and callbacks.
2. Call optional `RedSalamanderPluginShutdown()`.
3. Query optional `RedSalamanderPluginCanUnloadNow()`.
4. If it returns `FALSE`, keep the DLL mapped, keep the resource owner registered, mark the entry unload-deferred, skip same-path reload, and retry later.
5. `RedSalamanderPluginRetainModuleUntilProcessExit()` remains a process-shutdown-only escape hatch.

The backend command worker initializes COM as MTA (`COINIT_MULTITHREADED`) once for its entire lifetime; WPD operation and stream-reader entry points assert that this lifetime context is present instead of initializing and tearing COM down per command. Standalone picker workers initialize their own MTA. No detached thread may keep running through module unload without a reversible module keepalive owned by the quarantine/sweeper path.

Backend commands that can block in device I/O must run through a watchdog deadline. On timeout, the instance is marked disconnected, the user-facing operation returns `ERROR_DEVICE_NOT_CONNECTED`, `mtp.device.watchdog_trips` is emitted, and the still-running worker is moved to module-level quarantine instead of being blindly joined by the caller. The timed-out serialized device session is abandoned; a later `Initialize` creates a new worker, and every reader is bound to the backend generation that created it so a stale pre-timeout reader returns `ERROR_DEVICE_NOT_CONNECTED` without reaching the replacement queue. Live WPD sessions are recreated after abandonment. Copy/move progress and cancellation callbacks must run before entering the backend worker, and item-completion callbacks must run after the watchdog result returns on the caller thread; backend workers must not retain host callback or cookie pointers. Writer commit workers must own their staged byte buffer independently of the returned `IFileWriter` object so a timed-out commit cannot read caller-owned or writer-owned memory after `Commit` returns. `RedSalamanderPluginCanUnloadNow()` must return `FALSE` while a quarantined worker is live and may return `TRUE` only after the worker has completed and the sweeper has joined/erased it.

The current scaffold applies this watchdog/quarantine path to read, enumeration, attribute, basic-information, property-fetch, menu root enumeration, recursive directory-size helper calls, `CreateDirectory`, `CopyItem`, `MoveItem`, `DeleteItem`, `RenameItem`, and writer `Commit` entry points, including copy/move and writer overwrite verification. On timeout, the core queues `IMtpBackend::RequestCancel()` before quarantining the worker. The cancel request owns a plugin-module pin, contributes to `RedSalamanderPluginCanUnloadNow()`, resets its backend before finishing, and transfers the pin to `FreeLibraryWhenCallbackReturns(...)` as its final callback action so it cannot return through an unmapped DLL. The WPD backend tracks the active `IPortableDeviceContent` and active write stream under a mutex; cancellation calls `IPortableDeviceContent::Cancel()` and aborts the active stream with `IPortableDeviceDataStream::Cancel()` plus `IStream::Revert()`. The deterministic `mtp_watchdog_requests_backend_cancel` selftest guards the generic watchdog-to-backend cancel hook with a fake backend whose delayed operation only unblocks after `RequestCancel()` is issued. The deterministic fake-backend watchdog/cancel path is the v1 closeout gate; physical live-device proof for WPD cancellation/stream-release behavior is optional manual/post-closeout validation.

Runtime plugin refresh must respect that quarantine state. If the MTP module reports `CanUnloadNow()==FALSE`, `FileSystemPluginManager` must keep the old module in deferred-unload state, keep the resource owner registered, expose only a non-loadable unload-deferred placeholder in the public plugin list, skip same-path reload during rediscovery, and retry unload on a later refresh. After the worker exits and the sweeper reports unloadable, the next refresh may unload the deferred module and restore a fresh loadable plugin entry. The deterministic `mtp_runtime_refresh_defers_when_worker_quarantined` selftest guards this manager behavior.

`INavigationMenu::SetCallback(nullptr, nullptr)` is a synchronous quiet-point barrier for MTP navigation callbacks. The plugin must snapshot callback delivery state under lock, invoke `NavigationMenuRequestNavigate` outside that lock, track active callback delivery, and wait for active delivery to drain when clearing the callback. After the clear call returns, the previously registered callback and cookie must never be invoked again. The deterministic `mtp_unload_quiet_point_no_callback_after_clear` selftest guards this by blocking an in-flight navigation callback, proving clear waits, and then proving the same command cannot call the stale callback after the quiet point.

## Hot-Plug And Freshness

MTP v1 does not provide `IFileSystemDirectoryWatch`. Device hot-plug or WPD events invalidate device/session state; visible panes rely on manual refresh, host polling, or plugin-driven menu refresh. Device-gone conditions must return a specific failure instead of stale fabricated metadata.

Directory enumeration failures caused by a disconnected device must fail cleanly without returning a partial `IFilesInformation` object. The deterministic `mtp_disconnect_mid_enumeration_surfaces_error` selftest guards the fake-backend path by injecting one device-gone enumeration, requiring `ERROR_DEVICE_NOT_CONNECTED`, checking the out pointer stays null, and verifying the next enumeration can recover.

`mtp_live_device_smoke` is an opt-in production-backend smoke case. In default deterministic runs it must check `REDSALAMANDER_SELFTEST_MTP_DEVICE` before creating any WPD-backed plugin instance; if the variable is missing, it skips with an explicit reason rather than touching WPD or disappearing from the suite. When opted in, it creates a normal WPD-backed MTP plugin instance, enumerates the MTP root, selects the requested device or `*`, selects a writable storage root, switches the profile to `readOnly:false`, and runs only under an explicit scratch path. If `REDSALAMANDER_SELFTEST_MTP_SCRATCH` names a pre-existing non-empty scratch folder, the test must skip/fail safe rather than deleting unknown user data. The opted-in workflow exercises scratch create, file write/read, overwrite with `deviceReread`, same-folder rename, file copy, file move, file delete, and recursive scratch cleanup. The 2026-06-30 Full runner archive `Specs/TestRuns/feb0d5542efb/Continuation/2026-06-30_223700_mtp_full_green_handoff/full_last_run/` passed with `1125 total / 1078 passed / 0 failed / 47 skipped`; its corresponding Compare archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-30_214406/` records `mtp_live_device_smoke` as an explicit no-device skip because `REDSALAMANDER_SELFTEST_MTP_DEVICE` was not set. A 2026-07-01 read-only production WPD discovery probe is archived at `Specs/TestRuns/feb0d5542efb/Continuation/2026-06-30_223700_mtp_full_green_handoff/live_probe_no_wpd_2026-07-01/`; it ran the live-smoke case with an impossible requested device name and skipped with `MTP live smoke skipped: no WPD/MTP devices enumerated.` `Tools/Run-MtpLiveCloseout.ps1` is the preferred optional live-validation wrapper for both probe and live modes because it archives command/stdout/stderr/env/PnP/CIM/`last_run` evidence; its verified no-device probe is archived at `Specs/TestRuns/feb0d5542efb/Continuation/2026-06-30_223700_mtp_full_green_handoff/live_probe_helper_no_wpd_2026-07-01/` with `total=1 / passed=0 / failed=0 / skipped=1`. After the 2026-07-01 user-directed scope change, full physical-device execution, hotplug/device-removal behavior, WPD cancellation/stream-release proof, and live or recorded throughput evidence are optional post-closeout/manual validation, not v1 Done gates.

## Performance Validation

Metrics use the `mtp.*` namespace. Device-less fake-backend tests may validate round trips, call counts, byte counts, progress cadence, cancellation, and path normalization, but fake timing is not throughput evidence. The 2026-06-30 Compare archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-30_214406/perf/perf_metrics.jsonl` contains 647 `mtp.*` records across 40 unique metrics covering enum, props, transfer, writer, verify, path, overwrite, and device/watchdog families; audit summary: `Specs/TestRuns/feb0d5542efb/Continuation/2026-06-30_223700_mtp_full_green_handoff/mtp_metric_audit.md`. Production throughput or latency claims still require live-device smoke in an explicit scratch folder or recorded WPD replay evidence archived under `Specs/TestRuns/` before those numbers are cited, but that evidence is not a v1 Done gate after the 2026-07-01 scope change.

Directory enumeration must emit property-fetch counters. A batched WPD/fake-backend property path records `mtp.props.bulk_batches`; per-object fallback records `mtp.props.per_item_calls`. The deterministic `mtp_property_fetch_is_batched` selftest guards the expected fake-backend path with one batch and zero per-item calls.

Path identity handling must emit duplicate-name counters. Enumeration records `mtp.path.duplicate_groups` and `mtp.path.suffixed_entries`; failed ambiguous unsuffixed mutation records `mtp.path.ambiguous_resolve_failures`. The deterministic `mtp_duplicate_names_require_stable_suffix` selftest guards the fake-backend metric path with one duplicate group, two suffixed entries, and one fail-closed ambiguous mutation.

Overwrite verification must emit `mtp.verify.*` metrics. The deterministic `mtp_overwrite_byte_verify_level_matches_capability` selftest guards the configured path: `transmitHash` emits a local transmitted-byte hash with no fake-backend read, `deviceReread` emits re-read bytes after one fake-backend read, and `sizeOnly` emits the committed size only. Overwrite temp swaps emit `mtp.overwrite.temp_puid_present` after the non-empty temp PUID precondition passes and `mtp.overwrite.temp_swap_committed` after the temp is renamed into the final path; device-sourced copy/move overwrite swaps additionally emit `mtp.overwrite.device_source_temp_swap_committed`. PUID-less created-temp failures emit `mtp.overwrite.temp_puid_missing`, later policy-blocked overwrite attempts emit `mtp.overwrite.temp_puid_policy_blocked`, and failed original deletes after temp verification emit `mtp.overwrite.delete_original_failed`.

Overwrite journal paths must emit `mtp.overwrite.journal_intent_recorded` when the host-side intent file is written, `mtp.overwrite.journal_cleared` when a terminal success/rollback clears it, and `mtp.overwrite.journal_write_failed` when the step-0 intent write fails and the overwrite is aborted before device mutation. Temp create/copy failure paths must emit `mtp.overwrite.temp_upload_failed` for staged writer temp uploads and `mtp.overwrite.temp_copy_failed` for device-sourced copy/move temp creation. Replay paths must emit `mtp.overwrite.journal_replay_temp_committed` when a retained temp is renamed to the final path, `mtp.overwrite.journal_replay_temp_removed` when a retained temp is removed because the final path already exists, `mtp.overwrite.journal_replay_noop` when only a stale final-path journal remains, `mtp.overwrite.journal_replay_missing_objects` when neither final nor temp exists, `mtp.overwrite.journal_replay_failed` when recovery work fails, `mtp.overwrite.journal_replay_retry_scheduled` when recovery rename failure is retained for another bounded retry, `mtp.overwrite.journal_replay_rename_terminal` when recovery rename retries are exhausted and the journal is cleared while the temp remains, `mtp.overwrite.journal_replay_unavailable` when the host journal path cannot be opened and normal device commands continue, `mtp.overwrite.journal_replay_completed_swap_inferred` when a final-present/no-tempPUID replay is cleared by a destination-size match, `mtp.overwrite.journal_replay_stale_retained` when an unresolved retained replay rewrites the journal for another bounded attempt, `mtp.overwrite.journal_replay_stale_retained_terminal` when retained replay exhaustion is reached, `mtp.overwrite.journal_replay_stale_quarantined` when the stale journal is renamed to `.stale`, and `mtp.overwrite.journal_replay_stale_quarantine_failed` when that rename fails. Destination-present/temp-present cleanup delete failures emit `mtp.overwrite.journal_replay_failed`; a later successful cleanup retry emits `mtp.overwrite.journal_replay_temp_removed`. No-tempPUID orphan sweeps emit `mtp.overwrite.journal_orphan_sweep_temp_removed` for an exact single candidate, `mtp.overwrite.journal_orphan_sweep_retained` with the candidate count for zero or ambiguous retained entries, and `mtp.overwrite.journal_orphan_sweep_failed` on sweep or delete errors.

Device-sourced copy/move overwrite verification must emit source-kind metrics: `mtp.verify.device_source_trust_commit` for the default trusted-commit path, and `mtp.verify.device_source_bytes` plus `mtp.verify.device_source_reread_bytes` for the `deviceReread` opt-in path.

Transfer accounting must emit `mtp.transfer.*` metrics for deterministic byte counts. Successful fake-backend reads through a `CreateFileReader` reader emit `mtp.transfer.read_bytes` when `Read` transfers bytes; successful fake-backend copy emits `mtp.transfer.copy_bytes`; the fake-backend move partial-failure gate emits `mtp.transfer.move_fallback_bytes` with the copied byte count and the source-delete failure HRESULT. Production WPD whole-object uploads emit `mtp.transfer.write_bytes` after a successful stream commit.

Device-gone paths must emit `mtp.device.disconnects` when a disconnected device or injected disconnect is surfaced as a clean failure. The deterministic disconnect enumeration selftest guards one emitted count with the device-gone HRESULT.

Watchdog timeouts must emit `mtp.device.watchdog_trips`. The deterministic `mtp_hung_device_times_out` selftest guards the fake-backend read-path slice by injecting a delayed read, requiring prompt `ERROR_DEVICE_NOT_CONNECTED`, checking the instance remains device-gone until `Initialize`, proving a new command succeeds on the replacement worker, proving a stale pre-timeout reader fails immediately by generation, and proving `CanUnloadNow` stays false until the quarantined worker completes.

Backend cancellation requests issued by the watchdog must emit `mtp.device.cancel_requests` when the backend observes them. The deterministic `mtp_watchdog_requests_backend_cancel` selftest guards this path by requiring a delayed fake backend to unblock before its original delay after watchdog timeout.

The deterministic `mtp_mutating_create_directory_times_out` selftest guards the first mutating-command slice by injecting a delayed fake-backend `CreateDirectory`, requiring prompt `ERROR_DEVICE_NOT_CONNECTED`, checking the instance remains device-gone afterward, and proving unload stays deferred until the mutating worker completes.

The deterministic `mtp_mutating_item_commands_time_out` selftest guards item mutation commands by injecting delayed fake-backend `CopyItem`, `MoveItem`, `RenameItem`, and `DeleteItem` calls, requiring prompt `ERROR_DEVICE_NOT_CONNECTED`, checking the instance remains device-gone afterward, and proving unload stays deferred until each mutating worker completes.

The deterministic `mtp_writer_commit_times_out` selftest guards public writer commit by injecting a delayed fake-backend `WriteFile`, requiring prompt `ERROR_DEVICE_NOT_CONNECTED`, checking the instance remains device-gone afterward, and proving unload stays deferred until the command-owned staged-byte worker completes.

The deterministic `mtp_menu_and_directory_size_time_out` selftest guards metadata helpers by injecting delayed fake-backend menu root enumeration and recursive `GetDirectorySize` calls, requiring prompt return, checking the instance remains device-gone afterward, and proving unload stays deferred until each delayed worker completes.

Live write smoke tests must never delete or overwrite a pre-existing non-empty user folder.

## Unsupported In V1

Deferred or unsupported in v1:
- Native directory watch.
- Native indexed search.
- Alternate data streams.
- Random writes.
- Device-storage integrity claims without explicit `deviceReread`.
