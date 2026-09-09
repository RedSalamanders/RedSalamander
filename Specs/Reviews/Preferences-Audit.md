> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# Preferences Dialog Audit — Final Report

## 1. Verdict

The Preferences dialog is **not safe to ship as-is.** The single most dangerous theme is the **validation gap between the dialog and the settings loader**: five separate file-action editor paths (executable, action id, association primary action, extension/pattern match) let the user persist values that the loader's own `ValidateFileActionDefinition` / `IsValidFileActionId` / `IsValidFileActionExtension` reject — and on the next launch that one bad field triggers `RecoverSettingsLoadFailure`'s `out = Settings{}`, **wiping the entire settings store** (themes, shortcuts, plugins, folders, everything) to defaults. The dialog can write what the app cannot read back. The second theme is the **non-atomic two-file Apply state machine**: the main settings file is written to disk *before* the Monitor file, with no rollback and no baseline advance if the Monitor write fails — combined with an **unconditional re-save of stale in-memory settings on every close (including Cancel)**, this lets a Cancel either leak a partial apply to disk or overwrite an edit that already landed. Third is **security**: external-program file actions persist bare/relative executable names with empty working directories, and at launch `ShellExecuteExW` resolves them against the *browsed file's folder* — a textbook binary/DLL-planting path to code execution. Underneath these sit a cluster of **inverted/collapsed toggle bindings** (reduced-motion saved as its logical opposite; 4-mode display silently flattened to 2; ViewerWeb's security-relevant "Allow external nav"/"DevTools" toggles that can *never* persist the "on" value), **unconfirmed wholesale resets/imports** that nuke shortcut and file-action lists on one misclick, and an **uncommitted theme preview that leaks to disk on OK**. The root cause is structural: the dialog mutates a working copy with almost no field-level validation and relies entirely on a fragile, asymmetric commit path that does not mirror the loader's contract.

**Most dangerous, ranked:**
1. Unvalidated file-action fields → whole-store reset on next load (5 fields, 1 root cause).
2. Non-atomic two-file apply + unconditional close-save → partial apply / Cancel leaks / overwrites a landed edit.
3. Bare-exe + browsed-folder working dir → arbitrary code execution.
4. Inverted/collapsed toggle bindings → settings do the opposite of the label or can't be set at all.
5. Unconfirmed reset/import → one-click config loss.
6. Theme preview persisted on OK; theme name unbounded.

---

## 2. Findings by severity

### DATA-LOSS (whole-store reset / silent discard of a landed edit)

> **Shared root cause A — the file-action editor does not enforce the loader's contract.** Findings D1–D4 are the *same defect* surfacing through four fields: `SaveAction`/`SaveAssociation` validate only emptiness (and pluginId for the ViewerPlugin kind), never the charset/length/required-reference rules that `SettingsStore` enforces on load. Any violation makes `ParseFileActions` fail → `RecoverSettingsLoadFailure` runs `out = Settings{}` (`Common/Common/SettingsStore.cpp:5040`), resetting **the entire settings store** to defaults (a `.bak` of the bad file is kept, but the live store the user sees is wiped). Fix once: extract the loader's validators (`IsValidFileActionId`, `IsValidFileActionExtension`, the required-reference rule, the empty-executable rule) into a shared helper and call it from every file-action save path *before* mutating `workingSettings`, surfacing an `.rc` validation string and returning on failure.

**D1. Empty external-program executable persists an action the loader rejects → whole-store reset**
`RedSalamander/Preferences.FileActions.cpp:2236` — **data-loss, CONFIRMED.** Add External Program action, fill Action Id, leave Executable blank, Save → OK. `ValidateFileActionDefinition` returns `InvalidFileActionSettings` for empty `executablePath` (`SettingsStore.cpp:1826-1830`) → next launch resets everything. Mirror the empty-id/empty-pluginId early returns for the ExternalProgram branch.

**D2. Association primary action left on "(None)" persists empty `viewActionId`/`editActionId` → whole-store reset**
`RedSalamander/Preferences.FileActions.cpp:2068` — **data-loss, CONFIRMED.** The combo's index-0 item carries an empty value (`:1454`) and is the default selection (`:1470`). The serializer always writes `viewActionId` even when empty (`SettingsStore.cpp:4548`); the loader's `ValidateActionReference(..., required=true)` rejects it (`SettingsStore.cpp:1988`). Match `.log`, leave View on (None), Save → OK → store reset on next launch. Require a non-empty primary id before committing.

**D3. Action Id accepts characters/length the loader rejects → whole-store reset**
`RedSalamander/Preferences.FileActions.cpp:2205` — **data-loss, CONFIRMED.** `Trim` preserves internal spaces and Unicode; the loader's `IsValidFileActionId` (`SettingsStore.cpp:1571-1586`) requires `size <= 64`, leading ASCII-alnum, charset `[A-Za-z0-9_.-/]`. Ids like `my editor`, `_x`, or a 70-char id pass the dialog and reset the store on next load. Validate against the shared `IsValidFileActionId`.

**D4. Extension / pattern match values not validated against loader rules → whole-store reset**
`RedSalamander/Preferences.FileActions.cpp:2056` — **data-loss, CONFIRMED.** `BuildMatch`/`NormalizeExtensionText` only lowercase and prepend `.`; the loader's `IsValidFileActionExtension` (`SettingsStore.cpp:1590`) requires length 2..33 and a restricted charset. `.c++`, `.a b`, a bare `.`, or a non-ASCII extension persist and reset the store. Validate each extension/pattern (incl. `appliesTo` tokens) before commit.

**D5. Stale baseline after a failed Monitor save lets a later Cancel discard a value already written to the main file**
`RedSalamander/Preferences.Dialog.cpp:2710` — **config-loss (data-loss in effect), CONFIRMED.** This is the *consequence half* of root-cause B (below). After a partial apply (main written at `:2474`, monitor write fails), `CommitAndApply` early-returns at `:2723` without advancing `baselineSettings`. On Cancel, `RequestPreferencesDialogCancelClose` (`:2797`) only reverts the in-memory theme preview (`:1240-1242`) and never re-writes the main file — so the partially-applied main-file values persist permanently while the user believes they rolled back. Fix with a `mainFileWrittenSinceBaseline` flag that re-saves `baselineSettings` to the main file on Cancel, or roll the main file back at the point of monitor-save failure.

### CONFIG-LOSS

> **Shared root cause B — the two-file Apply is non-atomic and the close path re-saves unconditionally.** Findings D5 above and C1 here are the same machine: `SaveSettingsFromDialog` writes the **main** file first (`:2474`), advances `workingSettings` (`:2484`), *then* writes the separate **Monitor** file (`:2486`); on monitor failure `CommitAndApply` bails before advancing baseline / posting `kSettingsApplied` (`:2723-2741`), with **no rollback of the already-written main file**. Compounding it, `WM_NCDESTROY` calls `SaveSettingsAndSchema` **unconditionally on every close** (`:5418`), including a pure Cancel. Fix once: stage both files (serialize both, commit only if both writes succeed; write Monitor first or to temp), advance baseline only after all targets persist, and make the close-save write *only the window-placement subtree* when not dirty/applied — never the full document.

**C1. Partial apply: main file persisted while Monitor save fails, baseline left stale (and disk/memory diverge)**
`RedSalamander/Preferences.Dialog.cpp:2474` / `:2486` / `:2722` — **config-loss, CONFIRMED** (reported four times across panes; all the same root-cause B). Edit a main-pane field and a Monitor field, Apply with the Monitor file locked/read-only/disk-full: main file lands on disk, in-memory `*state.settings`/`baselineSettings` stay OLD, `kSettingsApplied` never posts, the dialog stays dirty, and the error names only the Monitor path. The Monitor edit is lost; disk and live state diverge until restart. See the staged-commit fix above.

**C2. Pure Cancel / no-op open still rewrites the whole settings file**
`RedSalamander/Preferences.Dialog.cpp:5413` — **config-loss, PLAUSIBLE.** Open Preferences, click Cancel immediately: `WM_NCDESTROY` unconditionally re-serializes the full settings document (`PrepareForSave` normalizes; hardcoded schemaVersion 16; **no unknown-key passthrough** per `Specs/Reviews/Settings-Persistence-Audit.md:32`). Any tolerated-but-unmodeled key (plugin-written, hand-added) is silently dropped on a gesture the user expects to be a no-op. The data loss is conditional on the no-passthrough defect, but the fix is the same: on a not-dirty/never-applied close, persist only the placement subtree.

**C3. "Open Monitor settings file" link materializes from stale baseline; later Apply clobbers external edits**
`RedSalamander/Preferences.Dialog.cpp:5087` — **config-loss, CONFIRMED.** With pending (unsaved) Monitor edits, clicking the link writes the not-yet-existing file from `monitorBaselineSettings` (`:5044`), not `workingMonitorSettings`. The user hand-edits that file; a later Apply overwrites it with the dialog's in-memory copy, and the external edits were never reconciled back in. Critically, the **hot-reload stale-guard does not cover the Monitor file** — the watcher is started only for `kAppId` (`RedSalamander.cpp:8938`), so `staleFromExternalReload` never fires and `ResolvePreferencesStaleSaveConflict` returns true with no prompt. Flush pending edits (or warn) and source on-create content from the working settings; treat the file as potentially stale after external editing.

**C4. Label-only hot-path slot is accepted in the UI but silently discarded on save**
`RedSalamander/Preferences.HotPaths.cpp:277` — **config-loss, CONFIRMED.** Give a slot a path (enables the Label field), type a label, blank the path. The UI keeps the slot (`MaybeResetWorkingHotPathsSettingsIfEmpty`, `:1006`) and marks dirty, but both serializer (`SettingsStore.cpp:6932-6938`) and parser (`SettingsStore.cpp:3430-3433`) gate the whole slot on a non-empty path — the label is written as `null` and gone on reopen. The edit appears to succeed and is lost with no warning. Clear the slot when the path goes empty, or persist label-only slots.

**C5. Reset-to-defaults wipes all keyboard shortcuts with no confirmation**
`RedSalamander/Preferences.Keyboard.cpp:2697` — **config-loss, CONFIRMED.** One click on "Reset to defaults" runs `emplace(CreateDefaultShortcuts())` immediately (guarded only by capture-inactive), discarding every custom binding in both scopes; persisted on OK/Apply. No `HOST_ALERT_QUESTION` anywhere in the file. Gate behind a modal Yes/No confirmation.

**C6. Import wholesale-replaces the entire shortcut list (and silently wipes the omitted scope)**
`RedSalamander/Preferences.Keyboard.cpp:3173` — **config-loss, CONFIRMED.** `shortcuts = std::move(imported)` with no confirmation/merge; `ParseShortcutsImportJson` zeroes both vectors first, so importing a file with only `functionBar` silently erases all `folderView` bindings. Confirm before replace, state it replaces (not merges), and preserve scopes absent from the file.

**C7. Reset Defaults wipes all viewer/editor associations and actions with no confirmation**
`RedSalamander/Preferences.FileActions.cpp:2186` — **config-loss, CONFIRMED.** Same missing-guard pattern as C5/C6: `ResetAssociationsAndActions` replaces the entire family with defaults on one click (no `Confirm`/`MessageBox` in the file), persisted on next Apply/OK. Modal-confirm before replacing; optionally scope to the selected entry.

**C8. Live theme preview persists to disk on OK even though it was never committed**
`RedSalamander/Preferences.Themes.cpp:1936` — **config-loss claimed; corrected to CONTRACT, PLAUSIBLE.** "Apply temporarily" mutates `*state.settings.theme` without `SetDirty` (`:1944-1948`); OK then takes the `!dirty` branch, skipping the validated `SaveSettingsFromDialog`/baseline-update path, and `WM_NCDESTROY` persists the previewed theme anyway (`:5418`). The Cancel path has a revert (`RestorePreviewAppliedPreferencesOnCancel`) but OK has none, and `previewApplied`/`baselineSettings` are left inconsistent. Reachability is narrow (selecting a real theme sets dirty; only the `ThemeSchemaSource::New` placeholder skips `SetDirty`), so this is a contract gap rather than clear high-severity loss — but the OK path should call the same revert (or not mutate shared settings for a transient preview). No `Apply-Temporarily → OK` self-test exists.

### SECURITY

> **Shared root cause C — external-program actions persist unrooted executables with no path validation, and the launcher resolves them against the browsed folder.** D1's empty-exe finding and the two security findings below share the same `SaveAction:2236`/`:2239` no-validation site; at runtime `LaunchExternalPlan` defaults `lpDirectory` to `CurrentDirectoryForContext` = the browsed item's parent (`FileActionLauncher.cpp:584/616`) and hands a bare `lpFile` to `ShellExecuteExW`. Fix at both ends: at config time require a fully-qualified, existing executable (reject bare/relative names); at launch, pass an absolute application name and never default the working directory to the browsed folder.

**S1 / S2. Bare/relative executable + empty working dir → binary/DLL-planting code execution**
`RedSalamander/Preferences.FileActions.cpp:2236` — **security, CONFIRMED** (reported twice; same root-cause C). Configure an external action with `notepad.exe` (or `..\evil.exe`, or a benign exe with a planted `version.dll`), empty Working Dir; view/edit a file inside an attacker-controlled folder (USB/download/network). `ShellExecuteExW` resolves the bare name and dependent DLLs through a search order that includes the attacker folder → arbitrary code execution as the user (CWE-426 untrusted search path). Cross-referenced to the confirmed launcher finding in `Specs/Reviews/AppShell-Audit.md:34-39`. Require absolute/existing executablePath at save; use `CreateProcessW` with an absolute app name and a trusted working dir at launch.

### CRASH / USE-AFTER-FREE

*(none confirmed at this severity — see "Coverage & residual risk")*

### VALIDATION

**V1. Import accepts `vk=0` / out-of-band vk, persisting a dead "VK_00" binding**
`RedSalamander/Preferences.Keyboard.cpp:3060` — **validation, CONFIRMED.** The only bound check is `vk > 0xFF || modifiers > 0x7` (`:3060`); `vk==0` with no modifiers passes, unlike the interactive `ApplyCapturedShortcut` which filters modifier-only/escape keys. `VkToStableName(0)` → "VK_00" round-trips through the store. The dead chord shadows the real default. Reject `vk<1` / non-real keys at import, reusing the capture-path filter.

**V2. Theme name unbounded on the edit path; only rejected later at export**
`RedSalamander/Preferences.Themes.cpp:2497` — **validation, CONFIRMED.** `nameEdit` handler does `def->name.assign(text)` with no length check; the 64-char contract is enforced only post-commit (`ThemeDefinitionIo.cpp:269/342`) and the main-store parser has no cap, so a >64 name persists to `settings.json` and later "Save theme to file" fails. Clamp to 64 in the change handler, mirroring `MakeUniqueUserThemeIdForRename`.

**V3. File-action editor: empty executable (non-runnable action) persisted with no feedback**
`RedSalamander/Preferences.FileActions.cpp:2206` — **validation, CONFIRMED.** Distinct from D1's reset cascade: even when the action doesn't crash the loader, an empty-exe ExternalProgram action is a dead config (runtime `BuildExternalLaunchPlan` returns `E_INVALIDARG`, View/Edit silently no-ops). Surface an "executable required" `.rc` string and return. (Folds into root-cause A's shared validator.)

**V4. File-action macro templates not syntax-validated at config time**
`RedSalamander/Preferences.FileActions.cpp:2238` — **validation, CONFIRMED.** `arguments`/`executablePath`/`workingDirectory` are stored raw; a malformed `{Filenam}` / stray brace / unknown `{Foo}` is accepted, and every later invocation fails to build a plan (a failure overlay appears at invocation, but no diagnosis at edit time). Run each template through a dry `ExpandMacros`/`TemplateReferencesMacro` at save and reject on invalid syntax/unknown macro.

**V5. Plugin numeric field clamp applies max after min → schema with min>max persists below the minimum**
`RedSalamander/Preferences.Plugin.Configuration.cpp:1762` — **validation, CONFIRMED.** `v = max(v,min); v = min(v,max)` — when `min>max` the second clamp wins, yielding `v==max < min`; no `min<=max` validation, and the empty-text branch writes `defaultInt` unconditionally. An untrusted schema (`min:10,max:5`) persists 5. Normalize to `lo=min(min,max), hi=max(min,max)`, validate the schema at parse time, and clamp the default the same way.

**V6. Plugin numeric field strips the minus sign on input → negative-min fields can't accept negatives**
`RedSalamander/Preferences.Plugin.Configuration.cpp:1089` — **validation, CONFIRMED.** `SanitizeIntegerText(text, false)` hard-codes `allowNegative=false`, dropping `-` on every keystroke and force-resetting the field, even when `field.hasMin && field.min < 0`. A `min:-100` field is locked to `[0,max]`. Pass `allowNegative = (field.hasMin && field.min < 0)`.

### INCORRECT-BEHAVIOR / CONTRACT

> **Shared root cause D — toggle/combo bindings invert or collapse the underlying enum.** Several panes render a multi-value enum as a 2-state Toggle (or map combo indices wrong), so the persisted value disagrees with the displayed label. These are independent sites but one reviewable pattern: any time a Toggle stands in for an N-value enum, the on/off choice indices must be derived from the schema, and round-trip self-tests must assert *label → persisted enum*, not dialog-to-dialog consistency.

**I1. Reduced-motion combo persists the inverted enum (On saved as Off, Off as On)**
`RedSalamander/Preferences.General.cpp:366` — **incorrect-behavior, CONFIRMED.** Selection handler maps index 1 ("On") → `ReducedMotionMode::Off` and index 2 → `On`; `SyncDxControlsFromState` mirrors the inversion so the dialog round-trips with itself, but every serializer/consumer reads canonically. Selecting "On" (reduce motion) writes `Off` → motion is **not** reduced — an accessibility setting does the exact opposite. The self-test encodes the same inversion and can't catch it. Swap the two mappings and add a `label→enum` assertion.

**I2. Pane display 2-state toggle silently collapses ExtraDetailed/Thumbnails to Detailed/Brief**
`RedSalamander/Preferences.Panes.cpp:181` — **config-loss in effect, CONFIRMED.** In non-high-contrast mode the 4-mode display is a 2-state Toggle; a Thumbnails/ExtraDetailed pane renders as unchecked ("Detailed"), and any toggle gesture overwrites the real mode with Brief/Detailed (persisted on OK). The 4-item combo already exists for the high-contrast path. Always use the combo, or render the combo when the mode is ExtraDetailed/Thumbnails.

**I3. ViewerWeb 2-option Toggle can only ever write the FIRST choice (toggleOn/Off both default to 0)**
`RedSalamander/Preferences.Plugin.Configuration.cpp:1205` — **incorrect-behavior, CONFIRMED — security-relevant.** `toggleOnChoiceIndex`/`toggleOffChoiceIndex` default to 0 and are **never assigned** in this file (unlike `ManagePluginsDialog.cpp:3282-3283`), so ON and OFF both write `choices[0].value`, and the display inverts for any default==`choices[1]`. ViewerWeb's "External navigation" (default "Allow") renders OFF and can **never** persist "Allow"; "DevTools" likewise. The user cannot enable these, and the first interaction silently rewrites to the wrong value. Compute the on/off indices from the schema and add a 2-option round-trip self-test.

**I4. Cache max-bytes round-trips bare byte values as KiB → 1024× inflation on blur**
`RedSalamander/Preferences.Internal.cpp:1412` — **incorrect-behavior, CONFIRMED.** `FormatCacheBytes` emits a bare digit string for non-1024-aligned values; `TryParseCacheBytes` defaults a unit-less number to KiB (`×1024`). A stored `1500` displays as "1500"; a no-op blur re-parses it to 1,536,000 and marks dirty → OK persists a 1024×-inflated cap; repeats compound. Reachable only for externally/legacy-set values. Make format/parse symmetric (emit/treat bare numbers as bytes) and add a `FormatCacheBytes∘TryParseCacheBytes == identity` self-test.

**I5. Custom bandwidth field treats empty/whitespace as 0 (Unlimited), silently dropping the throttle**
`RedSalamander/Preferences.FileOperations.cpp:161` — **incorrect-behavior, CONFIRMED.** In Custom mode, clearing the field parses as `0` and persists `defaultBandwidthLimitBytesPerSecond = 0` = Unlimited with no feedback; a user mid-retype unthrottles large copies over metered links. Distinguish "empty Custom field" (not-yet-entered, don't persist) from explicit Unlimited.

**I6. Active-FS disable guard uses case-sensitive compare while loader/quarantine use `EqualsNoCase`**
`RedSalamander/Preferences.Plugins.cpp:757` — **incorrect-behavior, PLAUSIBLE.** The only case-sensitive `currentFileSystemPluginId` compare among all consumers (others use `EqualsNoCase`). On a casing mismatch (hand-edited/migrated `settings.json5`) the guard is skipped, the in-use FS is disabled, and `ReloadPlugins` swaps the default FS and `FreeLibrary`s the module out from under panes still holding its `IFileSystem`. Requires an externally-introduced casing mismatch (the app writes canonical case), but it's a one-line fix to `EqualsNoCase` (and the debug guard at `:1286`).

**I7. Case-sensitive de-dup of custom plugin paths disagrees with the model/remove handler**
`RedSalamander/Preferences.Plugins.cpp:1967` — **incorrect-behavior, PLAUSIBLE.** Add-time `std::find` and save-time `sort/unique` are case-sensitive, so `C:\Plug\fs.dll` and `c:\plug\FS.DLL` both persist; but the grid's `FindRowIndexByPath` and the remove handler are case-insensitive, so Remove targets the wrong row. (The asserted loader duplicate-id cascade is **refuted** — the loader de-dups candidates case-insensitively at `FileSystemPluginManager.cpp:827`, so the DLL loads once; effect is a redundant row + confused Remove.) Use `EqualsNoCasePath` at add/save.

### LEAK / ROBUSTNESS

**R1. Plugin Test / Test-All synchronously `LoadLibrary` + run factory code on the UI thread**
`RedSalamander/Preferences.Plugins.cpp:1878` — **robustness, CONFIRMED.** "Test All" loops every plugin calling `TestPlugin` → `EnsureLoaded` (synchronous `LoadLibrary` + `RedSalamanderCreate`) on the UI thread, no worker/cancel. Network-backed plugins (Curl/S3/Drive) freeze the dialog for seconds; a hanging/malicious DLL's `DllMain`/factory hangs the UI thread indefinitely — violates the CLAUDE.md no-UI-blocking rule. Run probes on a background thread, marshal results via `PostDeferredAction`, show busy + cancel.

**R2. Adding a custom plugin path validates only `.dll` existence, then `LoadLibrary`s it unconditionally at startup**
`RedSalamander/Preferences.Plugins.cpp:1959` — claimed security; **corrected to ROBUSTNESS, PLAUSIBLE.** `IsDllPath` checks only existence+extension, not factory export/signature; the path is persisted and `LoadLibraryExW`'d at next launch (running `DllMain` before the export check). This is a deliberate "load a plugin DLL of your choice" gesture — a legitimate hardening gap (no signature/allow-list/warning), not an exploitable vuln absent prior code-exec. Probe the factory export at config time, warn that custom paths execute native code, prefer a managed dir/allow-list.

**R3. `browse:'folder'` plugin field persists an arbitrary, unvalidated path with no existence/type check**
`RedSalamander/Preferences.Plugin.Configuration.cpp:1162` — **robustness, CONFIRMED.** Typed text bypasses the dialog's `FOS_PATHMUSTEXIST`; `BuildConfigurationJson` writes `retainedText` raw with no `exists`/`is_directory` check (contrast the Value branch's clamp). A nonexistent/malformed/UNC path persists on blur. Validate at commit (`exists` + `is_directory`, reject empty/whitespace), surface an inline error, persist only on pass.

**R4. Editing a hot-path field re-runs `SyncDxControlsFromState` per keystroke, wiping undo/redo, selection, and trimming mid-edit**
`RedSalamander/Preferences.HotPaths.cpp:218` — **incorrect-behavior, CONFIRMED.** The per-keystroke `OnTextChanged` calls `SyncDxControlsFromState`, which unconditionally `SetText`s the focused field with the *trimmed* value; `TextField::SetText` clears undo/redo, drops the selection, and clamps the caret with no early-out. Ctrl+Z does nothing, trailing spaces vanish and the caret jumps, selections are lost mid-edit. Don't resync from the keystroke path (update model + SetDirty only; resync on blur/refresh), or skip `SetText` for the focused control / when text is unchanged.

**R5. `PrefsReorderPanelChildren` destroys panel children on the early-return path if any slot is null**
`RedSalamander/Preferences.Internal.cpp:157` — **robustness, PLAUSIBLE.** The helper `std::move`s wanted children out *before* the size-mismatch guard at `:157`; if `children` ever held a null slot, the guard returns after already nulling the moved-out slots, and the local `reordered` destroys the owned controls → corrupted/blank pane. Reachability is effectively nil today (the only insertion path `AddChild<>` never stores null, no per-slot reset exists, callers run synchronously), but the failure mode is destructive rather than a safe no-op. Validate inputs first and bail as a true no-op (compute the new order, commit atomically), so the early return never destroys controls.

---

## 3. Cross-cutting themes (highest-leverage fixes)

1. **One shared file-action validator, called before commit.** Root-cause A (D1–D4, V3) is the highest-leverage fix in the audit: a single helper that mirrors `IsValidFileActionId` / `IsValidFileActionExtension` / `ValidateActionReference` / the empty-executable rule, invoked from `SaveAction` and `SaveAssociation` before any `workingSettings` mutation, removes four data-loss vectors at once. The dialog must never be able to persist what the loader rejects. The deeper structural fix is that *no UI save path should be able to write a settings document the loader can't read back* — consider a "dry-load" of the serialized blob before Apply.

2. **Atomic two-file Apply + a close-save that only touches placement.** Root-cause B (C1, C2, D5) is one machine: stage both files, write Monitor-first or to temp, advance `baselineSettings`/post `kSettingsApplied` only after all writes succeed, roll back on partial failure, and make `WM_NCDESTROY` persist only the window-placement subtree when not dirty/applied. This simultaneously kills partial apply, the Cancel-overwrites-a-landed-edit case, and the no-op-Cancel full rewrite. Extend the hot-reload watcher to the Monitor app id so external-edit conflict detection (C3) actually fires for that file.

3. **Toggle-stands-in-for-enum binding discipline.** Root-cause D (I1, I2, I3) recurs because a Toggle is used for a multi-value enum and on/off indices are wrong or hard-collapsed. Standardize: derive on/off choice indices from the schema (as `ManagePluginsDialog` already does), render a combo when a Toggle can't represent the value space, and add round-trip self-tests that assert **label → persisted enum**, since dialog-to-dialog mirroring hides exactly these inversions.

4. **Field-level validation before persist, generally.** V2, V5, V6, R3, I4, I5 all share "the change handler commits whatever is typed; validation (if any) lives downstream." A consistent commit-time validate/clamp/normalize step per field — with inline error surfacing — covers theme name, plugin numeric min/max/negative, plugin folder paths, cache bytes, and bandwidth.

5. **Confirm destructive list operations; don't resync the focused editor.** C5/C6/C7 (wholesale reset/import with no modal) and R4 (per-keystroke resync clobbering edit state) are the two list-editor anti-patterns. Gate every wholesale replace behind a Yes/No `.rc` prompt; never `SetText` a focused field from a per-keystroke handler.

---

## 4. Coverage & residual risk

- **Severity recalibration.** Three findings were down-graded against their original labels: the Monitor "open settings file" data-loss claim and the theme-preview leak are **config-loss / contract** (not the worst case the title implies); the custom-plugin-path "add" finding is **robustness** (a deliberate plugin-load gesture, not an exploitable vuln); R2's de-dup cascade was **partially refuted** (the loader collapses case-variant candidates, so no duplicate-id conflict). The four file-action data-loss findings and both security findings are confirmed at full severity and must be fixed before release.

- **No confirmed crash / use-after-free.** The closest is the plugin case-sensitivity guard (I6), which *can* `FreeLibrary` an in-use FS module out from under a live `IFileSystem` — a real UAF mechanism, but gated on an externally-introduced casing mismatch (PLAUSIBLE, not confirmed reachable through in-app gestures). `PrefsReorderPanelChildren` (R5) is a latent destructive-on-error helper with effectively-nil current reachability. Both deserve the defensive fix even though neither is a live crash today.

- **Reachability caveats to re-test.** Several high-impact findings depend on externally-set or legacy values rather than pure in-app gestures: I4 (non-1024-aligned cache bytes), I6/I7 (casing mismatches in `settings.json5`), V5 (untrusted plugin schema `min>max`), C2 (tolerated-but-unmodeled keys). These are real but conditional; prioritize the validators anyway, since the dialog is precisely the place an externally-introduced bad value gets re-committed.

- **Residual risk / gaps not covered by this audit.**
  - **Persistence layer assumptions:** the whole data-loss class hinges on `RecoverSettingsLoadFailure`'s `out = Settings{}` and the no-unknown-key-passthrough writer (`Settings-Persistence-Audit.md:32`). If those behaviors change, severities shift; the file-action validators are the durable fix regardless.
  - **Launcher hardening** (S1/S2) is only half-fixable in the dialog; the `FileActionLauncher`/`ShellExecuteExW` working-directory behavior (`AppShell-Audit.md:34-39`) must be fixed in tandem.
  - **Self-test coverage is structurally weak for binding correctness:** at least three confirmed bugs (I1 reduced-motion, I3 ViewerWeb toggle, C8 preview-on-OK) are invisible to existing tests because the tests either encode the same inversion or only cover the Cancel path. Round-trip tests asserting persisted-value-equals-intended-value are missing across the dialog.
  - **Not audited here:** concurrency between the hot-reload watcher and an open dialog beyond C3; per-pane focus/keyboard-nav correctness; and the serializer/loader for panes other than file-actions, hot-paths, shortcuts, themes, plugins, and monitor.
