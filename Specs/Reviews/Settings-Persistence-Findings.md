> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# Settings-Persistence - verified findings (compact)

- **[config-loss/CONFIRMED]** `Common/Common/SettingsStore.cpp:5097` - Any non-16 schemaVersion (newer OR older build) backs up the file and wipes all settings to defaults
- **[config-loss/CONFIRMED]** `Common/Common/SettingsStore.cpp:5139` - A single invalid shortcut/fileAction/userMenu entry rejects and resets the ENTIRE settings store
- **[config-loss/CONFIRMED]** `Common/Common/SettingsStore.cpp:5139` - Single malformed shortcut/file-action/user-menu entry discards ALL user settings (connections, themes, hot-paths, folders) and renames the file away
- **[config-loss/CONFIRMED]** `Common/Common/SettingsStore.cpp:7344` - SaveSettings is an unguarded last-writer-wins full-blob overwrite with no stamp/version check at the store layer
- **[config-loss/CONFIRMED]** `Common/Common/SettingsStore.cpp:5046` - BackupBadSettingsFile MoveFileExW failure is non-fatal, so a corrupt file can be left in place while the app silently runs on defaults
- **[config-loss/CONFIRMED]** `Common/Common/SettingsStore.cpp:4232` - Non-ASCII/lone-surrogate string fields silently serialized as empty string, then dropped on reload (round-trip data loss)
- **[config-loss/CONFIRMED]** `Common/Common/SettingsStore.cpp:7344` - SaveSettings overwrites the whole store from a stale in-memory snapshot with no compare-and-swap (concurrent-writer lost update)
- **[config-loss/CONFIRMED]** `RedSalamander/Preferences.Dialog.cpp:2654` - Stale-save guard relies on asynchronous watcher delivery, leaving a TOCTOU window where a fresh external save is clobbered without the conflict prompt
- **[config-loss/CONFIRMED]** `Common/Common/SettingsStore.cpp:376` - No cross-process or in-process serialization of settings writes enables last-writer-wins between concurrent RedSalamander writers
- **[config-loss/CONFIRMED]** `RedSalamander/SettingsHotReload.cpp:147` - Main-app full-blob save has no stale-write detection: clobbers concurrent RedConfigure/ConnectionManager edits (connections, credentials, shortcuts)
- **[config-loss/CONFIRMED]** `RedSalamander/ConnectionManagerWindow.cpp:2663` - Connection secrets in Credential Manager are mutated before the profile list is persisted, with no rollback on save failure
- **[config-loss/PLAUSIBLE]** `Common/Common/SettingsStore.cpp:434` - Atomic write keeps no previous-good backup; a mid-write crash leaves only the .tmp and the just-overwritten target
- **[config-loss/PLAUSIBLE]** `Common/Common/SettingsStore.cpp:7344` - SaveSettings overwrites the entire settings blob from a pre-change in-memory snapshot with no store-level stale-write detection
- **[config-loss/PLAUSIBLE]** `RedConfigure/RedConfigureApp.cpp:84` - Workspace root resolution stops at the first ancestor containing .git, so RedConfigure can bind to the wrong repository/store
- **[config-loss/PLAUSIBLE]** `RedConfigure/RedConfigureSession.cpp:1119` - Default theme export filename collides across themes whose ids differ only by case or sanitized characters, silently overwriting a prior export
- **[corruption/CONFIRMED]** `RedSalamander/SettingsSchemaExport.cpp:423` - Duplicate plugin field keys produce a malformed JSON Schema with repeated object members
- **[corruption/CONFIRMED]** `Common/Common/ThemeDefinitionIo.cpp:448` - Theme save is non-atomic truncate-then-write; a crash or disk-full mid-write strands the user with a zero/partial theme file
- **[corruption/CONFIRMED]** `RedConfigure/RedConfigureSession.cpp:325` - Localization/theme export uses non-atomic truncate-in-place write (no temp+rename, no backup)
- **[corruption/CONFIRMED]** `RedConfigure/RedConfigureSession.cpp:337` - ofstream success is checked before close, so a failed flush on close is reported as a successful export
- **[corruption/CONFIRMED]** `RedConfigure/RedConfigureSession.cpp:347` - Theme export uses non-atomic truncate-then-write with no temp file or backup, so a crash/disk-full mid-write strands the user with a corrupt theme file
- **[credential-exposure/CONFIRMED]** `Common/Common/SettingsStore.cpp:3157` - Duplicate connection ids are not de-duplicated on load, colliding their Credential Manager secret entries
- **[credential-exposure/CONFIRMED]** `RedSalamander/ConnectionCredentialPromptDialog.cpp:593` - Unmasked secret can be copied to the global clipboard and is never cleared on dialog close
- **[credential-exposure/PLAUSIBLE]** `Common/Common/SettingsStore.cpp:3099` - Connection id loaded from settings is not validated for non-empty/uniqueness, allowing CredMan target-name collisions and silent secret loss
- **[credential-exposure/PLAUSIBLE]** `RedSalamander/ConnectionManagerWindow.cpp:2350` - Orphan-secret cleanup depends on an in-memory baseline that only covers removed ids, not changed/orphaned ones, and is skipped if save aborts
- **[credential-exposure/PLAUSIBLE]** `RedSalamander/ConnectionSecrets.cpp:73` - Secret in-memory wipe is best-effort and does not survive std::wstring reallocation/copies in the secret path
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/SettingsSchemaExport.cpp:253` - Integer schema bounds/defaults silently truncated to 32 bits via yyjson_get_int
- **[incorrect-behavior/CONFIRMED]** `RedConfigure/RedConfigureSession.cpp:1204` - Untranslated localization rows are exported with English source text as if translated
- **[incorrect-behavior/CONFIRMED]** `RedConfigure/Workspace/WorkspaceDiscovery.cpp:268` - Workspace directory scan silently stops on the first filesystem error mid-iteration, dropping all remaining themes/resource owners while still returning S_OK
- **[incorrect-behavior/PLAUSIBLE]** `RedConfigure/RedConfigureSession.cpp:1108` - Export root falls back to a relative path when workspace discovery yields an empty root
- **[migration-loss/CONFIRMED]** `Common/Common/SettingsStore.cpp:5097` - Strict `schemaVersion != 16` equality gate resets a newer (or older) file to defaults and renames it on round-trip through another app version
- **[migration-loss/CONFIRMED]** `Common/Common/SettingsStore.cpp:5317` - Any future/unknown settings key is dropped on save; cross-version round-trip destroys newer settings
- **[migration-loss/CONFIRMED]** `Common/Common/SettingsStore.cpp:5097` - Newer (or any non-16) schema version is treated as corrupt: startup load backs up and resets to defaults, destroying user settings on downgrade/round-trip
- **[migration-loss/PLAUSIBLE]** `RedConfigure/RedConfigureSession.cpp:1190` - Satellite .rc regeneration drops every existing translation whose resource id is no longer in the English source
- **[race/CONFIRMED]** `RedSalamander/SettingsHotReload.cpp:162` - Post-save file-stamp re-query races with concurrent external writer, marking the external change as already-applied and permanently dropping it
- **[robustness/CONFIRMED]** `Common/Common/SettingsStore.cpp:401` - WriteFileBytesAtomic uses a fixed `.tmp` name and keeps no previous-good copy; cross-process saves collide
- **[robustness/CONFIRMED]** `RedSalamander/SettingsHotReload.cpp:121` - Watcher re-arm gap can miss a settings change after FindNextChangeNotification failure
- **[robustness/CONFIRMED]** `RedSalamander/SettingsSchemaParser.cpp:146` - x-ui-order (and integer default) truncated from int64 to int without range check
- **[robustness/CONFIRMED]** `Common/Common/ThemeDefinitionIo.cpp:234` - Theme parser accepts unbounded input: whole file copied a second time with no size cap, enabling memory-exhaustion on a hostile theme file
- **[robustness/CONFIRMED]** `Common/Common/ThemeDefinitionIo.cpp:326` - A single malformed color value rejects the entire theme on load instead of degrading gracefully
- **[robustness/CONFIRMED]** `RedConfigure/RedConfigureSession.cpp:318` - Localization/theme export uses truncate-then-write (no atomic rename or backup), so a crash mid-write strands a truncated/empty .rc or theme file
- **[robustness/PLAUSIBLE]** `RedSalamander/SettingsHotReload.cpp:154` - Plugin schema sidecar write is not part of the atomic settings commit and is silently ignored on failure, producing a settings/schema mismatch window
- **[robustness/PLAUSIBLE]** `RedSalamander/SettingsSchemaParser.cpp:78` - Unsigned schema bound > INT64_MAX wraps to a negative value, corrupting validation range
- **[robustness/PLAUSIBLE]** `RedSalamander/SettingsSchemaExport.cpp:548` - Duplicate $defs / configurationByPluginId members on plugin-id sanitization+hash collision
- **[robustness/PLAUSIBLE]** `Common/Common/ThemeDefinitionIo.cpp:293` - Color map built from untrusted theme has no entry-count limit
- **[robustness/PLAUSIBLE]** `RedConfigure/RedConfigureSession.cpp:1148` - Review export collapses owners with colliding sanitized filenames, last writer wins
- **[validation/CONFIRMED]** `Common/Common/SettingsStore.cpp:3118` - Connection port deserialized and persisted without range validation (out-of-range value survives)
- **[validation/CONFIRMED]** `Common/Common/SettingsStore.cpp:3116` - Connection port and duplicate profile ids persisted without validation (out-of-range port, duplicate ids both retained)
- **[validation/CONFIRMED]** `Common/Common/SettingsStore.cpp:3116` - Connection TCP port persisted without range validation (out-of-range values survive round-trip)
- **[validation/CONFIRMED]** `RedSalamander/ConnectionManagerWindow.cpp:1957` - Connection port is persisted without range validation (0-65535), allowing invalid/truncated ports to be saved
- **[validation/CONFIRMED]** `Common/Common/ThemeDefinitionIo.cpp:126` - Duplicate top-level keys (id/name/baseThemeId) and duplicate color keys are silently resolved with no error, allowing a confusing/last-wins theme
- **[validation/CONFIRMED]** `RedConfigure/Localization/PlaceholderValidation.cpp:64` - Placeholder validation passes malformed brace syntax that std::vformat rejects, allowing a translation that never renders
