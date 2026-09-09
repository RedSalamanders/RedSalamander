> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# Preferences - verified findings (compact)

- **[config-loss/CONFIRMED]** `RedSalamander/Preferences.Panes.cpp:181` - Pane display 2-state toggle silently collapses ExtraDetailed/Thumbnails to Detailed/Brief and misrepresents the current mode
- **[config-loss/CONFIRMED]** `RedSalamander/Preferences.Dialog.cpp:2486` - Partial apply: main settings file is written to disk, but a failing Monitor-file save aborts the commit, leaving in-memory settings stale and the success silently dropped
- **[config-loss/CONFIRMED]** `RedSalamander/Preferences.Dialog.cpp:5087` - "Open Monitor settings file" link creates/opens the file from monitorBaselineSettings, ignoring pending dialog edits and risking later overwrite of external edits
- **[config-loss/CONFIRMED]** `RedSalamander/Preferences.Keyboard.cpp:2697` - Reset-to-defaults wholesale-replaces all user shortcuts with no confirmation
- **[config-loss/CONFIRMED]** `RedSalamander/Preferences.Keyboard.cpp:3173` - Import wholesale-replaces the entire shortcut list with no confirmation
- **[config-loss/CONFIRMED]** `RedSalamander/Preferences.FileActions.cpp:2186` - Reset-to-defaults wholesale-replaces all associations and actions with no confirmation
- **[config-loss/CONFIRMED]** `RedSalamander/Preferences.HotPaths.cpp:277` - Label-only hot-path slot (non-empty Label, empty Path) is silently discarded on save and cannot round-trip
- **[config-loss/CONFIRMED]** `RedSalamander/Preferences.Dialog.cpp:2474` - Partial apply: main settings file is persisted to disk before the Monitor file save, with no rollback if the Monitor save fails
- **[config-loss/CONFIRMED]** `RedSalamander/Preferences.Dialog.cpp:2710` - Stale in-memory baseline after a failed Monitor save lets a later Cancel discard a value already written to the main settings file
- **[config-loss/PLAUSIBLE]** `RedSalamander/Preferences.Dialog.cpp:5413` - On-close save persists settings even on a pure Cancel / no-op open (writes that should never happen)
- **[contract/PLAUSIBLE]** `RedSalamander/Preferences.Themes.cpp:1936` - Live theme preview persists to disk on OK even though it was never committed (uncommitted preview leak)
- **[data-loss/CONFIRMED]** `RedSalamander/Preferences.Dialog.cpp:2722` - Partial apply: successful main-file save is silently rolled back on disk when the Monitor file save fails
- **[data-loss/CONFIRMED]** `RedSalamander/Preferences.FileActions.cpp:2236` - Saving an external-program action with an empty executable path persists a value the loader rejects, wiping the entire settings store on next launch
- **[data-loss/CONFIRMED]** `RedSalamander/Preferences.FileActions.cpp:2068` - Saving a viewer/editor association with the primary action set to (None) persists an empty viewActionId/editActionId, which fails reload and resets the whole store
- **[data-loss/CONFIRMED]** `RedSalamander/Preferences.FileActions.cpp:2205` - Action Id field accepts characters/length the loader rejects, persisting an id that resets the whole store on next load
- **[data-loss/CONFIRMED]** `RedSalamander/Preferences.FileActions.cpp:2056` - Extension and pattern match values are not validated against loader rules, allowing a persisted match that resets the whole store
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/Preferences.Dialog.cpp:2486` - Partial apply: main settings file persisted to disk while monitor save fails, leaving in-memory baseline un-advanced
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/Preferences.General.cpp:366` - Reduced-motion combo persists the INVERTED enum value ("On" saved as Off, "Off" saved as On)
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/Preferences.Internal.cpp:1412` - Cache max-bytes field round-trips bare byte values as KiB, inflating the stored size by 1024x on blur
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/Preferences.FileOperations.cpp:161` - Custom bandwidth field accepts an empty/whitespace value as 0 (Unlimited), silently dropping a throttle the user intended to set
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/Preferences.Plugin.Configuration.cpp:1205` - Two-option 'option' field rendered as a Toggle can only ever read/write the FIRST choice (toggleOn/Off both default to 0)
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/Preferences.HotPaths.cpp:218` - Editing a hot-path Path/Label field re-runs SyncDxControlsFromState on every keystroke, calling SetText on the focused field and wiping its undo/redo history, selection, and trimming text mid-edit
- **[incorrect-behavior/PLAUSIBLE]** `RedSalamander/Preferences.Plugins.cpp:757` - Active-filesystem disable guard uses case-sensitive compare while the loader/quarantine use EqualsNoCase — the in-use FS plugin can be disabled and unloaded
- **[incorrect-behavior/PLAUSIBLE]** `RedSalamander/Preferences.Plugins.cpp:1967` - Duplicate custom plugin paths differing only by case are added and persisted (case-sensitive de-dup mismatch)
- **[robustness/CONFIRMED]** `RedSalamander/Preferences.Plugins.cpp:1878` - Test / Test-All synchronously LoadLibrary and run plugin factory code on the UI thread
- **[robustness/CONFIRMED]** `RedSalamander/Preferences.Plugin.Configuration.cpp:1162` - browse:'folder' Text field persists an arbitrary, unvalidated path with no existence/type check at config time
- **[robustness/PLAUSIBLE]** `RedSalamander/Preferences.Plugins.cpp:1959` - Custom plugin path add accepts any existing .dll with no plugin validation and persists it for unconditional LoadLibrary at startup
- **[robustness/PLAUSIBLE]** `RedSalamander/Preferences.Internal.cpp:157` - PrefsReorderPanelChildren destroys all panel children if any child slot is null (moved-from leak on early return)
- **[security/CONFIRMED]** `RedSalamander/Preferences.FileActions.cpp:2236` - File-action editor persists a bare/relative executable name with no absolute-path validation (binary/DLL-planting launch from the browsed untrusted folder)
- **[security/CONFIRMED]** `RedSalamander/Preferences.FileActions.cpp:2236` - External-program action permits a bare (non-rooted) executable name with empty working directory, enabling binary planting from an untrusted browsed folder
- **[validation/CONFIRMED]** `RedSalamander/Preferences.Keyboard.cpp:3060` - Import accepts vk=0 and raw out-of-band vk values, persisting a dead/VK_00 binding
- **[validation/CONFIRMED]** `RedSalamander/Preferences.Themes.cpp:2497` - Name edit writes unbounded theme name into settings, only rejected later at file export
- **[validation/CONFIRMED]** `RedSalamander/Preferences.FileActions.cpp:2206` - External-Program action saved with an empty executable path (invalid, non-runnable action persisted)
- **[validation/CONFIRMED]** `RedSalamander/Preferences.FileActions.cpp:2238` - Argument / executable / working-directory macro templates are not syntax-validated at config time
- **[validation/CONFIRMED]** `RedSalamander/Preferences.Plugin.Configuration.cpp:1762` - Value field min/max clamp applies max after min, so an (untrusted) schema with min>max persists a value below the declared minimum
- **[validation/CONFIRMED]** `RedSalamander/Preferences.Plugin.Configuration.cpp:1089` - Numeric 'value' field strips the minus sign during input, so a field with a negative min cannot accept any negative value
