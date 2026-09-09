> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# RedSalamander — whole-application deep-audit: consolidated summary

10 area audits · **306 confirmed + 133 plausible** verified defects (582 candidates, 143 refuted at the adversarial-verify stage). Per-area reports in this folder; index in `_AuditProgress.md`.

## App-wide cross-cutting themes (ranked by leverage)

These patterns recur across many subsystems — fixing the *pattern* beats fixing instances.

1. **Non-atomic / in-place destructive writes that lose user data.** Truncate-then-write or overwrite-in-place with no temp+atomic-rename and no previous-good copy. Instances: same-FS overwrite via `CopyFileExW` CREATE_ALWAYS (FS); `Win32FileWriter` CREATE_ALWAYS+delete-on-abort (FS); SettingsStore whole-store reset + non-atomic 2-file apply (Settings/Preferences); `LocalSearchIndexCore` snapshot save; image export; Monitor Save-As. **The correct template already exists** (the cross-FS bridge's temp-stage→verify→atomic-promote) — apply it everywhere.

2. **Trusting unvalidated sizes / counts / offsets from untrusted input.** Header/field/stream-derived values flow into allocations, indices, and EOF decisions with no bound. Instances: unknown-size COPY treats a 0-byte read as EOF → silent truncation (FS bridge, 7z, S3, Curl); viewer parsers (PE tables, RAW/WIC dimensions, EXIF IFD); Monitor ETW property lengths; FolderView OLE/clipboard buffers walked with no `GlobalSize`; DxUi virtualization indices vs mid-frame data change; host trusts plugin `count`. **Validate against the real buffer/size before allocate/index; tighten the `IFileReader::Read` contract.**

3. **Validation gap between the writer and the reader.** The UI/dialog persists values the loader later rejects → whole-config wipe. Instances: Preferences file-action fields (5) → `Settings{}` reset; out-of-range TCP port; vk:0 shortcut; unvalidated theme id. **The writer must enforce the reader's contract — share one validator.**

4. **Enable-state / UI-gate is the only guard; the real invocation path bypasses it.** Destructive commands reachable via accelerator/`SendMessage(WM_COMMAND)` with no precondition re-check (Commands); drag bypasses menu-grey (FolderView); disabled-plugin re-enable via posted WM_COMMAND. **Re-validate preconditions at the handler, not just in the menu.**

5. **TOCTOU / stale-snapshot acting destructively.** A target/selection/enumeration captured at one time, acted on at another. Instances: command selection/target across modal pumps; FolderView drag path list + `_itemsFolder`; FS cross-FS move deletes from a stale enumeration; settings concurrent-writer clobber. **Re-resolve by stable key / re-verify before the irreversible step.**

6. **Use-after-free from async/callbacks after teardown, and unenforced thread-affinity.** Instances: DxUi post-callback dangling `Control*` (dispatcher + 5 controls) and cross-thread UIA UAF; Monitor worker capturing raw `self` + Document's lock-defeating reference accessors; viewer decode/teardown races; FolderView icon worker + `_itemsFolder` race; HostServices secret-map races. **Own worker lifetime (cleanup groups), re-validate after re-entrant callbacks, don't return references that outlive the lock, marshal to the UI thread structurally.**

7. **Synchronous blocking I/O on the UI thread** (CLAUDE.md violation). Instances: NavigationView (5 paths: GetFileAttributes/GetVolumeInformation/GetDiskFreeSpace/icons/plugin-mount on offline UNC); viewers (decode, libvlc load, PE join on remote read); Monitor; commands (`WNetGetUniversalName`, HKCR walk, recursive `remove_all`); Preferences (plugin/theme enum). **Move to a worker + post results back; the codebase already has the pattern.**

8. **Security: DLL/binary planting, ShellExecute injection, web sandbox, no impersonation, plaintext secrets.** Instances: launcher inherits attacker CWD + no `SetDefaultDllDirectories`; file-action runs a bare exe from the browsed folder; `cmd /K pushd` lets cmd `%VAR%`-expand; ViewerWeb is an unsandboxed `file://` HTML host with default-allow navigation (NetNTLM/exfil); search service runs as LocalSystem with no client impersonation; credential clipboard reveal + plaintext-at-rest concerns.

9. **Reparse-point following with no cycle/depth guard.** FS traversal/live-scan (stack overflow), metadata copy stamps the link target, ViewerSpace scan. **Skip/visited-set reparse points; open with FILE_FLAG_OPEN_REPARSE_POINT.**

10. **Errors swallowed / incompleteness reported as success.** FS search silently drops entries with no warning flag; `GetDirectorySize` swallows access-denied; Monitor/clipboard failure paths. **Add a partial/incomplete channel and surface it.**

## Prioritized fix roadmap

**P0 — data-loss on ordinary gestures (fix first):**
- FS: symmetric unknown-size COPY verification; stage+atomic same-FS overwrite (FS-Subsystem #1, #2).
- Settings: section-scoped load recovery (stop whole-store reset on one bad entry) + validate-on-write; atomic 2-file Apply (Settings #1–3, Preferences D1–D5/C1).
- Preferences: shared file-action validator enforcing the loader contract before commit.
- Commands: confirmation/staging for Unpack overwrite, makeFileList truncate, Pack self-delete; inline-rename `..` reject.
- FolderView: same-folder/self/descendant drop guard + plumb the drop point; report MOVE only after the copy verifies.
- App shell: `WM_QUERYENDSESSION`/`WM_ENDSESSION` save-on-exit + drain in-flight ops.

**P0 — security:**
- App shell: `SetDefaultDllDirectories(SYSTEM32|APPLICATION_DIR)` in launcher + main; pass install dir as child CWD; require fully-qualified file-action executables.
- ViewerWeb: render untrusted HTML in a sandboxed synthetic origin with deny-by-default CSP + full navigation mediation; default external-nav to block.
- Search service: impersonate the named-pipe client before enumerating.

**P1 — crash / UAF / hangs:**
- DxUi: re-validate `target` via `ControlBelongsToTree` after every dispatched handler (project item B-S0-1); fix UIA cross-thread access.
- Monitor: bounded scrollback + bounded queue (drop-oldest); thread-pool cleanup group; fix Document reference-returning accessors.
- Viewers: parser bounds (pixel/dimension caps, ETW-style length checks) + teardown-race guards.
- Move UI-thread blocking I/O to workers across NavigationView, viewers, Monitor, commands, Preferences.
- FS: reparse-point cycle guard in traversal.

**P2 — robustness / validation / leaks:** per-pane value validation, error surfacing, handle/COM/temp-file cleanup, the long robustness tails in each report.

## Residual risk (not statically resolvable — needs a human / harness)
- Backend-dependent truncation (SFTP/SIZE-less-FTP, MinIO/Ceph S3) — kill-mid-transfer harness.
- Crash/power-loss atomicity of overwrites & saves — kill-the-process harness.
- IME/TSF + UIA + D3D device-removed — interactive/stress testing.
- ViewerWeb exploit chains — live WebView2 red-team.
- A **hostile-`IFileSystem` plugin harness** (pathological `ReadDirectoryInfo` / `RedSalamanderEnumeratePlugins`) exercises the FS, ViewerSpace, and host `count`-OOB classes at once.
- Crafted-file fuzz corpus for every parser (PE/RAW/EXIF/TIFF/archive/ETW) under ASan/UBSan.
