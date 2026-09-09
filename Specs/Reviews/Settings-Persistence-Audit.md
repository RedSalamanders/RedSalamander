> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


I'll produce the final audit report directly. The findings are already verified with evidence and line numbers, so I'll synthesize, deduplicate by root cause, and prioritize them.

# Settings/Config/Credential Persistence Audit — Final Report

## 1. Verdict

**RedSalamander's persistence layer is structurally unsafe for any multi-writer or version-skewed deployment, and it loses user data on inputs that are not even corrupt.** The single settings store (`Common/Common/SettingsStore.cpp`) is the blast radius for almost every serious finding. Four themes dominate and account for nearly all confirmed config-loss:

- **Over-eager "recovery" destroys good data.** A version-integer mismatch (`schemaVersion != 16`) or a *single* malformed shortcut/file-action/user-menu entry renames the entire real settings file to `.bad.<timestamp>` and boots on factory defaults — wiping all connections, themes, hot-paths, shortcuts, and history. No genuine JSON corruption is required. A version downgrade or a future-version round-trip is sufficient.
- **No concurrency control at the store layer.** `SaveSettings` is an unconditional last-writer-wins full-blob overwrite with no stamp/CAS check and no cross-process lock. The main app, RedConfigure, and the hot-reload watcher all write the same file; whichever saves last silently clobbers the others' connections, credentials, and shortcuts. The only stale-write guard lives in one UI dialog and depends on async watcher timing (TOCTOU).
- **Credential integrity rides on an unvalidated profile `id`.** Duplicate connection ids are never de-duplicated, so two profiles collide on the same Credential Manager target — one authenticates with another's secret, and deleting one destroys the other's. Credential Manager is mutated *before* the profile list is persisted with no rollback, so a failed save permanently desynchronizes the two stores.
- **No backup, no atomicity beyond torn-write protection.** The "atomic" write protects against torn files but keeps *no previous-good copy*; a logically-bad save (e.g. defaults persisted after a transient load failure) permanently overwrites real data with nothing to roll back to. Theme and RedConfigure exports are worse — plain truncate-then-write with no temp file at all.

Plus one real **credential-exposure**: a revealed password in the credential prompt can be copied to the global clipboard and is never cleared on close.

Treat the SettingsStore load/recovery and save/concurrency rewrite as the top priority; everything else is downstream of it.

---

## 2. Findings by Severity

### TIER 1 — Config-loss / Corruption / Credential-exposure / Crash

---

**Strict `schemaVersion != 16` equality gate backs up the file and resets to defaults on any other version (newer OR older)**
`Common/Common/SettingsStore.cpp:5097`
Severity: **config-loss / migration-loss** — CONFIRMED.
*(Deduplicated: this single root cause was reported as four separate findings — store-write, store-read, hotreload, and a no-unknown-keys variant. They are the same bug and its serialization mirror.)*

- **Failure scenario:** A newer build writes `schemaVersion 17` (or any v2–v15 file exists); an older/parallel/rolled-back build launches. `LoadSettingsFromResolvedPath` hits `if (schemaVersion != 16)` → `RecoverSettingsLoadFailure(UnsupportedSchemaVersion)` with `backupBadFile=true, fallbackToDefaults=true` (line 5241). The valid file is `MoveFileExW`-renamed to `<name>.bad.<timestamp>` (line 313) and the app boots on `Settings{}` (line 5040). The `>= 2` (5122) and `>= 5` (3996) branches are dead — nothing but 16 ever reaches them. There is no migration ladder.
- **Companion data-loss (same root cause):** `SaveSettings` rebuilds the whole document from the typed struct and hardcodes `schemaVersion 16` (line 5332); the loader parses only modeled keys with no passthrough container. Two builds at version 16 with different key sets silently drop each other's unknown keys on every round-trip.
- **Deeper fix:** Replace equality with a supported-range + forward-migration ladder. Accept `<=16` and migrate up. For strictly-newer versions, **refuse to overwrite** — load read-only or surface an explicit "newer version" error *without renaming the file*, so a downgrade can never strand the user's real config in an unseen `.bad`. Persist unmodeled top-level keys via a whole-file "extra" passthrough so same-version cross-build round-trips don't drop newer settings. Reserve backup+reset for genuinely unparseable JSON only.

---

**A single malformed shortcut / file-action / user-menu entry discards the ENTIRE store and renames the file away**
`Common/Common/SettingsStore.cpp:5139` (and `:5127`, `:5133`)
Severity: **config-loss** — CONFIRMED.
*(Deduplicated: reported twice — store-write and store-read — same mechanism.)*

- **Failure scenario:** One shortcut with an out-of-range `vk`, a `commandId` not starting with `cmd/`, or a wrong JSON type makes `ParseShortcuts` return `ERROR_INVALID_DATA` on the *first* bad binding (lines 4112–4123). `LoadSettingsFromResolvedPath` routes that to `RecoverSettingsLoadFailure` → `out = Settings{}` (5040) + `BackupBadSettingsFile` (5048). Every saved connection, theme, hot-path, and folder entry is renamed to `.bad.*` and the app comes up on defaults. `ParseFileActions` (5127) and `ParseUserMenuSettings` (5133) behave identically.
- **Inconsistency:** `ParseConnections`, `ParseFolders`, `ParseTheme`, `ParseHotPaths` are `void` and `continue` past bad records — so connections/folders/themes can't nuke the store, but three sections can. The discard policy is incoherent.
- **Deeper fix:** Parse each top-level section into its own recover scope; on a section error, keep all successfully-parsed sections and reset *only* the failing section (or skip the single bad entry with `Debug::Warning`), matching the resilient `void` parsers. Reserve whole-file discard for unparseable JSON / wrong root.

---

**`SaveSettings` is an unguarded last-writer-wins full-blob overwrite — concurrent writers silently clobber each other (and the hot-reload self-stamp suppresses the dropped change)**
`Common/Common/SettingsStore.cpp:7344` / `:5301`; `RedSalamander/SettingsHotReload.cpp:147`,`:162`
Severity: **config-loss / race** — CONFIRMED.
*(Deduplicated: reported five times — store-read, store-serialize, two store-concurrency, hotreload — all the same missing compare-and-swap at the store layer. Kept as one because the fix is one.)*

- **Failure scenario:** RedConfigure (separate process) adds an S3 connection profile + credential reference and saves. Before the main app's watcher reloads, the user assigns a hot-path (`RedSalamander.cpp:6507`) or exits (window placement save, `:2566`). `SavePreparedSettingsAndSchema` serializes the *pre-edit* in-memory snapshot and `MoveFileExW(MOVEFILE_REPLACE_EXISTING)` (line 434) replaces the whole file — deleting the new profile. The orphaned CredMan secret is stranded. Worse: it then calls `MarkAppliedStamp` (`SettingsHotReload.cpp:162`), so the pending `kSettingsFileChanged` sees `ShouldIgnoreStampLocked == true` (lines 65–69) and the external edit is silently dropped — *no conflict alert ever fires*.
- **Post-write stamp race (same family):** The applied stamp is re-`stat`'d in a separate syscall after the write (`SettingsHotReload.cpp:162–166`), not captured atomically. If an external writer replaces the file between the write and the `stat`, the *foreign* file's stamp is stored as applied, permanently masking that external change as already-applied (`SettingsStore.cpp:5167` opens a fresh handle).
- `SettingsFileStamp` infrastructure already exists (`TryGetSettingsFileStampForPath:5167`) but is used only on the reload path, never to guard writes. The only stale-write detection lives in the Preferences UI flag (`staleFromExternalReload`), which RedConfigure and main-app programmatic saves bypass entirely.
- **Deeper fix:** Add a store-layer compare-and-swap: capture `SettingsFileStamp` at load, and in a `SaveSettingsIfUnchanged` variant re-read the on-disk stamp immediately before the `MoveFileExW` and fail with a distinct HRESULT if it differs, forcing merge-and-retry. Have `WriteFileBytesAtomic`/`SaveSettings` *return* the committed bytes' own stamp (queried on the final handle) instead of re-`stat`-ing. Wrap load→mutate→save in a cross-process named mutex (`Global\RedSalamander.Settings.<AppId>`) and re-read+merge before serializing. Route *all* writers (main app, RedConfigure, programmatic) through this guarded path.

---

**Stale-save guard is asynchronous-only — TOCTOU window clobbers a fresh external save with no conflict prompt**
`RedSalamander/Preferences.Dialog.cpp:2654`
Severity: **config-loss** — CONFIRMED.

- **Failure scenario:** An external editor saves the settings file a few ms before the user clicks Apply, *or* the dialog isn't marked dirty when the change arrives. `ResolvePreferencesStaleSaveConflict` only prompts when `state.staleFromExternalReload` is true (2656), which is set *only* inside the async `OnSettingsReloadedFromDisk` (2701) and *only* if `state.dirty` (2689). If the user clicks Apply before that message pumps, the resolver returns true and the stale `merged = *state.settings` blob overwrites the external change. No `PromptStaleSaveConflict` shows.
- **Deeper fix:** Make the check synchronous at save time: in `CommitAndApply`, call `TryGetSettingsFileStamp` and compare against the stamp recorded at last disk load; raise the conflict prompt on mismatch regardless of watcher-message timing. Do not depend on the message pump for data-loss prevention.

---

**Atomic write keeps no previous-good backup; a logically-bad save permanently overwrites real data**
`Common/Common/SettingsStore.cpp:434` / `:376`
Severity: **config-loss** — CONFIRMED (mechanism), with the crash-during-write sub-case correctly capped (MoveFileEx protects torn writes).

- **Failure scenario:** A transient load failure earlier in the session left the app on defaults (`out = Settings{}`, and a *read* failure passes `backupBadFile=false` at line 5066, so no `.bad` is even made). The user changes one preference; `WriteFileBytesAtomic` atomically replaces the good file with the near-empty defaults blob. No prior copy exists to roll back. All connections/themes/hot-paths gone, no `.bad` to recover from.
- The fixed `.tmp` name (`path + ".tmp"`, opened share-mode `0`) also collides across concurrent same-appId writers (the second `CreateFileW` fails `ERROR_SHARING_VIOLATION` and silently loses that save).
- **Deeper fix:** Before the `MoveFileExW`, rotate the existing target to a single-slot `<name>.bak` (use `ReplaceFileW`'s `lpBackupFileName`, which also preserves ACLs, or explicit copy-to-`.bak`). On load, if the primary is missing/corrupt, try `.bak` before defaults. Guard against persisting a freshly-defaulted snapshot over a previously non-default file. Use a PID/GUID-unique temp name.

---

**`BackupBadSettingsFile` failure is non-fatal — corrupt file left in place, then overwritten by defaults with no copy aside**
`Common/Common/SettingsStore.cpp:5046`
Severity: **config-loss** — CONFIRMED.

- **Failure scenario:** A corrupt file is detected but the `.bad` rename fails (file held open by the watcher/AV, read-only dir, or 100 timestamp collisions). `RecoverSettingsLoadFailure` returns `S_FALSE` regardless (the `std::nullopt` is logged at 316 and dropped); `recovery.backedUp` stays false. The caller only warns when `backedUp` is true (`RedSalamander.cpp:7405`, no else-branch), so the user sees nothing. The corrupt-but-real original remains at the live path; the next save `WriteFileBytesAtomic`-overwrites it with in-memory defaults. The salvageable original is destroyed, never copied aside.
- **Deeper fix:** Treat backup failure as significant: if `BackupBadSettingsFile` returns `nullopt`, do **not** reuse the original path for subsequent saves — retry with sharing-tolerant copy semantics, or surface a hard error so silent default-then-overwrite is blocked. Combine with the save-side `.bak`.

---

**Non-ASCII / lone-surrogate string fields silently serialized as `""`, then dropped on reload (round-trip data loss)**
`Common/Common/SettingsStore.cpp:4232` (`NewString`), `:116` (`Utf8FromUtf16`)
Severity: **config-loss** — CONFIRMED.

- **Failure scenario:** A connection name/host/initialPath, hot-path slot, or shortcut contains a lone UTF-16 surrogate (legal in NTFS paths/filenames). `Utf8FromUtf16` passes `WC_ERR_INVALID_CHARS`, so `WideCharToMultiByte` fails and returns `""`. `NewString` detects `utf8.empty() && !value.empty()` and substitutes JSON `""` (4235), `AddStringObjectMember` returns `S_OK`, the whole save returns `S_OK`. On reload, `ParseHotPaths` (3430) discards empty-path slots and `ParseConnections` (3157) drops profiles with empty id/name — the connection and its CredMan secret association silently vanish. Valid astral-plane pairs round-trip fine; only ill-formed sequences trigger it.
- **Deeper fix:** `NewString` must not coerce a non-empty wide string to `""`. Either propagate failure (return `nullptr` → caller aborts the save, preserving the good on-disk file) or drop `WC_ERR_INVALID_CHARS` and emit best-effort/replacement encoding so original bytes survive. Add a self-test round-tripping a `\xD800` path through connection name and hot-path.

---

**Duplicate connection ids are not de-duplicated on load → Credential Manager target collision (wrong-secret auth + secret destruction)**
`Common/Common/SettingsStore.cpp:3157` / `:3099`; `RedSalamander/ConnectionSecrets.cpp:43`
Severity: **credential-exposure** — CONFIRMED.
*(Deduplicated: reported three times — store-serialize, store-read validation, secrets. Same root cause: no `id` uniqueness. The empty-id sub-claim from the secrets variant is REFUTED — line 3157 already drops empty-id profiles.)*

- **Failure scenario:** Two profiles share a non-empty `id` (hand-edit, import, cross-machine merge, copy/paste of a profile block) and both pass the line-3157 emptiness filter. `ParseConnections` de-duplicates *names* (3168–3204) but has no `usedIds` set. `BuildCredentialTargetName` (`ConnectionSecrets.cpp:43`) derives the CredMan target purely from the id, so both map to the same target `.../<id>/password`. Profile B authenticates with A's stored secret; `DeleteGenericCredential` → `CredDeleteW` (`ConnectionSecrets.cpp:178`) on B destroys A's still-referenced secret. The secret-access authorization map (`ConnectionSecrets.cpp:404/415/426`) is also id-keyed, so authorization state cross-contaminates too.
- Note: ids are normally GUIDs (`CoCreateGuid`), so collisions only arise from externally-supplied/tampered input — but the consequence is real credential cross-contamination, not mere validation.
- **Deeper fix:** In `ParseConnections`, track seen ids in an `unordered_set`; on collision regenerate a fresh GUID (leaving the canonical profile's secret intact) or skip the duplicate, *before* any id-keyed credential is read/written. Enforce id uniqueness symmetrically on persist. Make `BuildCredentialTargetName` empty-result a hard error at call sites rather than a silent no-op.

---

**Connection secrets mutated in Credential Manager BEFORE the profile list is persisted, with no rollback on save failure (cross-store desync)**
`RedSalamander/ConnectionManagerWindow.cpp:2663`
Severity: **config-loss / credential-exposure** — CONFIRMED.

- **Failure scenario:** User deletes a profile and changes another's password, clicks Save. `SaveConnectionsSettings` runs `DeleteSecretsForRemovedConnections` (2663) and the `CommitSecretsForProfile` loop (2665–2682) — irreversible `CredWriteW`/`CredDeleteW` — *first*, then `SaveSettingsAndSchema` (2689). If the save fails (disk full, file locked by watcher/RedConfigure, power loss), it returns false (2693) with no credential restored. CredMan has lost the removed profile's secret (or written the new password) while `connections.json` still lists the old set. The two non-transactional stores are permanently inconsistent; the user must re-enter passwords with no indication why.
- **Deeper fix:** Persist the settings file **first** (atomic temp+rename), and only on success mutate Credential Manager. Or snapshot prior credential values before delete/overwrite and restore them on save failure. Stage all CredMan mutations and apply only after the file write succeeds.

---

**Revealed secret can be copied to the global clipboard and is never cleared on dialog close**
`RedSalamander/ConnectionCredentialPromptDialog.cpp:593`; `Common/DxUi/DxUi.TextInput.cpp:2490`
Severity: **credential-exposure** — CONFIRMED.

- **Failure scenario:** User unmasks the password (Show toggle, `SetMasked(false)` at 596), selects it, and copies (Ctrl+C/X, Shift+Insert, Ctrl+Insert — all route through `OnCopy`). `TextField::OnCopy` only blocks copy while `_masked` (2490); once unmasked it writes plaintext via `CopyTextToClipboard`. The peek/reveal state also bypasses the guard. `CopyTextToClipboard` sets `CF_UNICODETEXT` with no `ExcludeClipboardContentFromMonitorProcessing`/`CanIncludeInClipboardHistory` formats. Confirm/Cancel/`WM_NCDESTROY` `SecureClear` the in-memory buffers but never `EmptyClipboard`. The plaintext lingers on the system clipboard and is captured into Windows Clipboard History / cloud clipboard, potentially syncing to other devices. (Self-initiated reveal+copy, not a silent leak.)
- **Deeper fix:** Gate `TextField::OnCopy` on an explicit "sensitive" flag set by the prompt, not the transient `_masked`/reveal state. When copying a revealed secret, set the clipboard-history-exclusion formats. On teardown, if this dialog last wrote the clipboard, `EmptyClipboard` it.

---

**Theme save and RedConfigure exports are non-atomic truncate-then-write with no temp file, rename, or backup**
`Common/Common/ThemeDefinitionIo.cpp:448` → `RedSalamander/Preferences.Internal.cpp:1099`; `RedConfigure/RedConfigureSession.cpp:325`,`:347`
Severity: **corruption** — CONFIRMED.
*(Deduplicated: theme save + three RedConfigure export variants — same `std::ofstream(path, trunc)` anti-pattern.)*

- **Failure scenario:** User overwrites an existing `user/forest` theme (or re-exports an `.rc`/`.theme.json5`). `TryWriteFileFromString` / `WriteUtf16LeTextFile` / `WriteBinaryFile` open `std::ofstream(path, std::ios::binary)` which truncates the live file to zero *before* writing. A crash/power-loss/disk-full between truncate and flush leaves a zero-byte or half-written file. On next load `ParseThemeDefinitionJson5` returns `EmptyInput`/`ParseFailed` and the hand-tuned theme is gone — no `.bad`, no backup. The settings store already does this correctly (`SettingsStore.cpp:434` + `BackupBadSettingsFile:310`); these paths use none of it.
- **Companion (same files):** `output.good()` is checked *before* the destructor flushes on close (`RedConfigureSession.cpp:337`,`:357`), so a close-time flush failure (disk fills mid-export) is reported as a *successful* export — the UI shows "Export done" over a truncated file.
- **Deeper fix:** Write to a sibling temp, `FlushFileBuffers`, then `ReplaceFileW`/`MoveFileExW(MOVEFILE_REPLACE_EXISTING|WRITE_THROUGH)`, keeping a `.bak`. Explicitly `flush()`+`close()` and re-test `good()`/`!fail()` before returning success.

---

**Duplicate plugin field keys produce a malformed JSON Schema with repeated object members**
`RedSalamander/SettingsSchemaExport.cpp:423`
Severity: **corruption** — CONFIRMED.

- **Failure scenario:** A plugin's `GetConfigurationSchema()` returns a `fields` array with two entries sharing a `key` (buggy/hostile plugin, copy/paste). `BuildPluginConfigJsonSchema` calls `yyjson_mut_obj_add(properties, keyOut, prop)` (423) per field — append-only, no existing-key check. The plugin-ID level *is* de-duplicated (`addedIds` set, 515/524/574) but the per-field loop is not. The aggregated `<appId>.settings.schema.json` ends up with a repeated member name under `properties` — malformed/ambiguous JSON. Strict validators / JSON5 readers / RedConfigure tooling see invalid or last-key-wins schema; schema-driven validation of the user's real plugin config becomes wrong.
- **Deeper fix:** Track emitted field keys in an `unordered_set<string_view>` and `continue` on repeat, or use `yyjson_mut_obj_put` (replace) semantics. Mirror the plugin-id de-dup already present.

---

### TIER 2 — Migration-loss / Race / Validation

---

**Satellite `.rc` regeneration drops every existing translation whose id is no longer in the English source**
`RedConfigure/RedConfigureSession.cpp:1190`; `RcWriter.cpp:41`
Severity: **migration-loss** — PLAUSIBLE (loss requires manual copy-back of the regenerated `.rc`).

- **Failure scenario:** A resource id is *renamed* in the English `.rc` but the satellite still holds a valid human translation under the old id. `MergeStringTables` iterates only source (target is a lookup-by-source-id map), so target-only ids are never emitted; `BuildSatelliteRcStringTable` writes a fresh table without them. Export targets a `RedConfigureOutput` staging dir, so loss only occurs if the user copies the regenerated file back over the live per-language `.rc` — at which point all drifted-out translations are silently gone. No "N entries will be dropped" warning exists.
- **Deeper fix:** Union source ids with ids already present in the parsed target satellite; carry forward (and flag as orphaned, don't delete) target-only translations, or explicitly warn how many entries will be dropped.

---

**Untranslated localization rows exported as English source text masquerading as real translations**
`RedConfigure/RedConfigureSession.cpp:1204`
Severity: **incorrect-behavior** — CONFIRMED.

- **Failure scenario:** Editing one cell makes a whole owner/culture "dirty" (`ReviewOwnerCultureHasDirtyCell`, 1185). The export loop then emits *every* row's `targetText` (1204), and untranslated cells were seeded with `targetText = sourceText` (544/1059). The generated satellite `.rc` contains English text for untouched strings, marked as if genuinely translated. Re-merging masks the genuinely-missing translations (they now look complete), degrading the dataset.
- **Deeper fix:** Export only cells with `hasExistingTranslation || dirty`, or emit untranslated entries in a clearly-untranslated form.

---

**Placeholder validation passes malformed brace syntax that `std::vformat` later rejects (translation never renders)**
`RedConfigure/Localization/PlaceholderValidation.cpp:64`
Severity: **validation** — CONFIRMED.

- **Failure scenario:** A translator enters `Fichier {0} ouvert {` or `{foo`. `ExtractPlaceholders` treats an unbalanced/non-digit brace as plain text (`continue` at 64–67, 78–81) and returns `Ok`. `UpdateTranslation` accepts and exports it. At runtime `FormatStringResource` → `std::vformat` throws `format_error` on the stray `{`; the catch (`Helpers.h:603`) logs and falls back to English. The translated string silently never shows, though RedConfigure reported it valid.
- **Deeper fix:** Reject target text with unbalanced braces or `{`/`}` that aren't valid placeholders/escapes (`{{`,`}}`). Run a trial `std::vformat` with stub args during validation and surface any `format_error`.

---

**Watcher re-arm gap can miss a settings change after `FindNextChangeNotification` failure**
`RedSalamander/SettingsHotReload.cpp:121`
Severity: **race / robustness** — CONFIRMED (narrow; failure-path only, change delayed not lost).

- **Failure scenario:** On `FindNextChangeNotification` failure the handle is dropped (121–125) and re-armed via `FindFirstChangeNotificationW` (87), which only reports changes *after* it returns. A write landing in that gap is never signaled; with no periodic stamp poll, the app runs on stale settings until an unrelated directory event wakes the watcher.
- **Deeper fix:** After re-arming, immediately post one `kSettingsFileChanged` (or directly compare `TryGetSettingsFileStamp` vs `lastAppliedStamp`); the stamp dedup makes a spurious post harmless.

---

**Connection TCP port persisted without range validation (out-of-range survives round-trip)**
`Common/Common/SettingsStore.cpp:3116` / `:6618`; `RedSalamander/ConnectionManagerWindow.cpp:1957`
Severity: **validation** — CONFIRMED.
*(Deduplicated: reported four times — same missing 1..65535 clamp on parse, persist, and UI input.)*

- **Failure scenario:** A port like `99999` (hand-edit/import, or typed into the UI where `wcstoul` can saturate to `ULONG_MAX` then truncate to uint32) is accepted with no clamp (`ParseConnections:3116`, `TryValidateAndNormalizeConnectionProfiles` validates only the name). It's re-persisted verbatim (6618) and formatted into the authority `host:99999` (`ConnectionProfileUtils.cpp:237`), producing an opaque connect failure. The FileSystemCurl `>65535` checks validate only the plugin's own URL parse, not the typed `profile.port`.
- **Deeper fix:** Validate on UI input, on parse, and before persist: reject or clamp to 1..65535 (treat 0/out-of-range as protocol default), surface a warning when coercing.

---

**Orphan-secret cleanup relies on a fragile in-memory baseline; external reload / cross-writer removal leaves CredMan secrets orphaned forever**
`RedSalamander/ConnectionManagerWindow.cpp:2350`
Severity: **credential-exposure** — PLAUSIBLE (orphaned secret at rest under user's own DPAPI, not disclosure; save-abort sub-claim REFUTED — deletion runs before the save).

- **Failure scenario:** `DeleteSecretsForRemovedConnections` only deletes targets for ids in `_baselineConnectionIds` absent from the current set, and the baseline is rebuilt from disk on every external hot-reload (`LoadConnections:2957`). If RedConfigure removes profile X from disk while the manager is open, the reload drops X from both `_connections` and the baseline, and the delete never runs — X's password/refresh-token stays in Credential Manager indefinitely. Any id removed while the manager was closed is likewise never reconciled.
- **Deeper fix:** Enumerate actual CredMan targets under the `RedSalamander/Connections/` prefix (`CredEnumerateW`) and delete any whose id is absent from the current persisted profile set, instead of diffing a single-session baseline.

---

**Duplicate top-level/color keys in theme files silently resolved first-wins (scalars) vs last-wins (colors)**
`Common/Common/ThemeDefinitionIo.cpp:126`
Severity: **validation** — CONFIRMED.

- **Failure scenario:** A hostile/hand-edited theme has two `id:` entries. `ReadRequiredString` uses `yyjson_obj_get` → first occurrence (126); the colors loop uses `parsed.colors[key] = argb` → last occurrence (331). A reviewer reading the file sees one value while the app stores another; on re-save only one is emitted, silently rewriting the file to a different effective theme than reviewed.
- **Deeper fix:** Reject documents with duplicate object keys (root and colors) with a clear `InvalidData` message.

---

### TIER 3 — Incorrect-behavior / Contract / Robustness

---

**A single malformed color value rejects the entire theme on load**
`Common/Common/ThemeDefinitionIo.cpp:326`
Severity: **robustness** — CONFIRMED. One bad hex (`#fff`, `red`, `#GGG`, trailing-space `#FF0000 `) makes the colors loop `return InvalidData` for the whole document (326–328); `TryParseColorUtf8` accepts only exact `#RRGGBB`/`#AARRGGBB`. A carefully-built imported theme fails entirely over one typo. (Refuses to load; does not destroy the on-disk file.) **Fix:** skip the offending key, collect a warning, keep valid colors; hard-fail only on structural errors.

---

**Theme parser copies the whole file a second time with no size cap (memory-exhaustion / `std::terminate` on hostile theme)**
`Common/Common/ThemeDefinitionIo.cpp:234`
Severity: **robustness** — CONFIRMED. The import reader (`Preferences.Internal.cpp:1069`) has no size cap (unlike `SettingsStore`'s 16 MiB `kMaxSettingsFileBytes`); `ParseThemeDefinitionJson5` then does `std::string mutableJson(jsonText)` — a full second copy — inside a `noexcept` function, so `bad_alloc` under memory pressure calls `std::terminate` (hard crash, losing unsaved Preferences edits). **Fix:** add a `kMaxThemeFileBytes` cap at the reader and parser top; parse `YYJSON_READ_INSITU` to avoid the second copy.

---

**Color map from untrusted theme has no entry-count limit**
`Common/Common/ThemeDefinitionIo.cpp:293`
Severity: **robustness** — PLAUSIBLE (the 100 MB scenario is bounded by the 16 MiB read cap, but ~1M tiny entries within the cap still bloat the in-memory map and re-serialized settings). **Fix:** cap accepted color entries (e.g. 256) or validate against a known key allow-list.

---

**Plugin schema sidecar write is separate from the settings commit and swallowed on failure**
`RedSalamander/SettingsHotReload.cpp:154`
Severity: **robustness** — PLAUSIBLE (corruption harm REFUTED — the per-app `.settings.schema.json` is never read back by the load/coercion path; it's an editor IntelliSense artifact, self-healing on next save). Real gap: schema-write failure is logged and swallowed while `saveHr` is returned as success. **Fix:** propagate the failure or regenerate the schema deterministically on load rather than persisting a second artifact.

---

**Duplicate `$defs` members on plugin-id sanitize+hash collision**
`RedSalamander/SettingsSchemaExport.cpp:548`
Severity: **robustness** — PLAUSIBLE (export-only; requires a sanitizer collision *plus* a full 32-bit FNV-1a collision across the handful of built-in plugins — astronomically unlikely; the property-key half of the claim is REFUTED). **Fix:** de-dup on the generated `defName`, not the raw id; disambiguate on collision.

---

**Integer schema bounds/defaults truncated to 32 bits via `yyjson_get_int`; `x-ui-order` narrowed to `int`; UINT > INT64_MAX wraps negative**
`RedSalamander/SettingsSchemaExport.cpp:253` / `SettingsSchemaParser.cpp:146`,`:78`
Severity: **incorrect-behavior / robustness** — CONFIRMED (export fidelity at `SchemaExport.cpp:253` and `displayOrder` narrowing at `SchemaParser.cpp:146`); the UINT-wrap at `SchemaParser.cpp:78` is PLAUSIBLE (latent — the schema parser's bounds have no production consumer; only the test project reads them). **Fix:** use `yyjson_get_sint`/`_uint` + `add_sint`/`add_uint`; range-check before narrowing; add a `min<=max` sanity check; store `displayOrder` as `int64_t` or `std::clamp`.

---

**Workspace directory scan silently stops on the first filesystem error mid-walk, returning S_OK with a partial catalog**
`RedConfigure/Workspace/WorkspaceDiscovery.cpp:268` / `:297`
Severity: **incorrect-behavior** — CONFIRMED. `for (; !ec && it != end; it.increment(ec))` with only `skip_permission_denied`: a long-path, dangling reparse point, or transient lock sets `ec` and terminates the loop (also via in-loop `is_directory(ec)`/`is_regular_file(ec)`), dropping all later themes/owners. `DiscoverWorkspace` returns `S_OK` regardless; truncation is invisible. **Fix:** record the failing path into `outResult.errors`, clear `ec`, and continue; `\\?\`-prefix the root.

---

**Workspace root resolution accepts the first ancestor with `.git`, can bind to the wrong repository**
`RedConfigure/RedConfigureApp.cpp:84`
Severity: **config-loss** — PLAUSIBLE (needs the exe deployed inside an unrelated git checkout with no closer `RedSalamander.sln`; field is user-editable for recovery). The `|| exists(.git)` disjunct with no preference ordering makes the nearest unrelated repo win, so exports land in the wrong tree. **Fix:** scan all ancestors for `RedSalamander.sln` first; fall back to `.git` only if no `.sln` exists anywhere.

---

**Export root falls back to a relative path when workspace discovery yields an empty root**
`RedConfigure/RedConfigureSession.cpp:1108` / `:1119`
Severity: **incorrect-behavior** — PLAUSIBLE (the named "launched from unrelated dir" trigger is inaccurate — startup uses absolute exe-dir; empty root is only reached by clearing the workspace field). With an empty root, exports resolve relative to the process CWD. **Fix:** validate `_workspace.root` is non-empty and absolute before composing paths; fall back to a well-defined absolute location and surface it in the UI.

---

**Review export / default theme export filename collisions silently overwrite a prior export**
`RedConfigure/RedConfigureSession.cpp:1148` (review) / `:1119` (theme)
Severity: **robustness / config-loss** — PLAUSIBLE (review-export trigger REFUTED — owner names are `.vcxproj` stems and can't contain filename-illegal chars; theme trigger requires two user-theme ids collapsing to the same case-folded/sanitized stem). `SanitizePathPart` doesn't guarantee uniqueness and `WriteBinaryFile` truncate-overwrites with no existence check. **Fix:** track emitted destinations; append a numeric suffix or prompt before clobbering a differing existing file; use case-folded uniqueness across the catalog.

---

## 3. Cross-cutting Themes

1. **"Recovery" is the data-loss vector.** Every load-path safety mechanism — version gate, per-section parse, `.bad` backup — destroys the *only good copy*. The `.bad` rename is the sole "protection" and it's a rename-away of the user's real data, often followed by an immediate defaults-overwrite. The whole recovery design needs inverting: never move/discard the original until a verified copy exists, and degrade per-section instead of per-file.

2. **One settings file, three uncoordinated writers, zero concurrency control.** Main app, RedConfigure, and the hot-reload watcher share `<appId>.settings.json` with no cross-process lock, no compare-and-swap, and a stale-write guard that lives in one UI dialog and depends on async message timing. Atomicity is misunderstood as concurrency safety — `MoveFileEx` prevents *torn* files but does nothing about *lost updates*. This is the highest-leverage fix: a store-layer stamped CAS + named mutex closes ~6 of the confirmed config-loss findings at once.

3. **Identity and bounds validation are absent on externally-sourced data.** Profile `id` (credential key), `port`, color keys, schema integers, duplicate JSON keys — none are validated/clamped/de-duplicated on the trust boundary. The id gap is the most dangerous because it silently maps to credential storage.

4. **Atomic-write discipline is inconsistent across the codebase.** `SettingsStore` does temp+flush+`MoveFileEx` (but no `.bak`); theme save and *all* RedConfigure exports do bare `ofstream(trunc)`. The good pattern exists and should be factored into one shared helper (temp + flush + `ReplaceFileW` with `.bak`, returning the committed stamp) and used everywhere.

5. **Secrets touch plaintext at the edges.** CredMan/DPAPI protects at-rest, but secrets are mutated before the profile write (no rollback), copied to the global clipboard when revealed, orphaned when profiles vanish via reload, and held in `std::wstring` whose wipe is best-effort. Defense-in-depth around the secret lifecycle is missing.

6. **Failures are silently swallowed.** Schema-sidecar write failure, `ofstream` close-flush failure, backup-rename failure, partial directory scan — all report success. Callers can't react to losses they're never told about.

---

## 4. Coverage & Residual Risk

**Needs crash/fault injection (cannot be confirmed by static review alone):**
- The save-side "no `.bak`" and "defaults-overwrite-good-file" loss (`SettingsStore.cpp:434`) under real power-loss/disk-full timing — verify the exact window and that `.bad` truly never fires on a valid-but-empty overwrite.
- Theme/RedConfigure truncate-then-write loss under crash mid-write — confirm zero-byte/partial result and the false-success on close-flush failure (`RedConfigureSession.cpp:337/357`).
- The `noexcept` theme-parser `std::terminate` on `bad_alloc` (`ThemeDefinitionIo.cpp:234`) — confirm hard crash vs graceful failure under a large file.

**Needs cross-process / multi-tool concurrency testing (a human running two writers):**
- Main-app ↔ RedConfigure ↔ watcher last-writer-wins clobber and the `MarkAppliedStamp` self-suppression that hides the dropped change — needs a deterministic two-process harness racing on the same file.
- The post-write stamp re-stat race (`SettingsHotReload.cpp:162`) and the Preferences TOCTOU window (`Preferences.Dialog.cpp:2654`) — both are timing-dependent and won't surface in single-threaded tests.
- Watcher re-arm gap (`SettingsHotReload.cpp:121`) — requires forcing `FindNextChangeNotification` failure.

**Needs hostile/edge-input fixtures (recommend adding to self-tests):**
- Lone-surrogate (`\xD800`) round-trip through connection name, host, initialPath, hot-path, shortcut — the most insidious silent loss.
- Duplicate connection ids → CredMan target collision (wrong-secret auth + secret destruction) — a true credential-integrity test, ideally exercising the real CredMan write/delete.
- `schemaVersion` 15 and 17 files → confirm the `.bad` rename + defaults reset, and drive the future migration ladder once implemented.
- Single malformed shortcut/file-action/user-menu entry among valid data → confirm whole-store discard, then per-section isolation after fix.

**Needs human judgment / product decision:**
- Localization regeneration loss (`RedConfigureSession.cpp:1190`) and untranslated-as-translated export (`:1204`) — whether to warn, carry forward orphans, or block; depends on the intended translator workflow.
- Clipboard exposure (`ConnectionCredentialPromptDialog.cpp:593`) — acceptable risk tolerance for a self-initiated reveal+copy vs. enforcing no-copy on sensitive fields.
- Schema-export 32-bit truncation / UINT-wrap — low live impact today (no production consumer of the parser bounds), but a latent trap if a future feature wires the parsed bounds into validation; decide whether to fix now or gate behind that feature.

**Residual blind spots not covered by this audit:** the actual DPAPI/CredMan blob handling under roaming-profile sync; settings file ACLs/permissions on multi-user machines; behavior when the settings directory itself is read-only or on a network share; and whether any plugin persists its own secrets outside this store (out of scope here — only the FileSystemCurl port-parse path was sampled).
