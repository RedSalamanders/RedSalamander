---
name: build-preferences-pages
description: Create or update RedSalamander settings-backed Preferences pages without breaking dirty tracking, Apply/OK persistence, Cancel semantics, schema metadata, localization, accessibility, or tests. Use whenever adding a Preferences category/control or changing how a Preferences control reads, previews, commits, or serializes a setting.
---

# Build Preferences Pages

Read `Specs/UI/UI_PreferencesDialog.md`, `Specs/Core/Core_SettingsStore.md`, and the affected page implementation before editing.

## Follow the complete settings path

Treat every control as one end-to-end path:

```text
Settings model -> parse/write/prune -> schema -> workingSettings
-> dirty projection -> save merge projection -> Apply/OK -> runtime notification
```

Do not consider a control complete because it renders or mutates `workingSettings`.

1. Add or update the typed setting in `Common/SettingsStore.h`.
2. Add parsing and writing in `Common/Common/SettingsStore.cpp`.
3. Normalize default optional sections in `RedSalamander/SettingsSave.h`.
4. Add stable UI metadata to `Specs/SettingsStore.schema.json`.
5. Add localized labels/descriptions to the embedded and satellite RC files.
6. Build the page on the shared DxUi page host and expose distinct accessible names/help text.
7. In every user callback:
   - ignore programmatic synchronization,
   - update only `state.workingSettings` (or the page's documented working document),
   - call `SetDirty(dialog, state)` with the actual root Preferences dialog HWND,
   - apply only documented live preview behavior,
   - resynchronize controls from the draft.
8. Extend both manual projections in `Preferences.Dialog.cpp`:
   - `IsDirty(...)` must compare the setting's effective baseline and working values;
   - `SaveSettingsFromDialog(...)` must merge the changed working section into `merged`.
9. Keep `Cancel` draft-only. After successful `Apply`/`OK`, advance the baseline and notify settings consumers through the existing commit pipeline.
10. Update the authoritative Preferences and settings-store specs.

## Preserve default semantics

Compare effective values, not merely `std::optional::has_value()`, when a missing section means defaults. When saving a changed section, assign the working optional section so reverting to defaults can remove it. Ensure the writer and `SettingsSave::PrepareForSave(...)` omit default-only sections without losing explicit non-default values.

Never update `baselineSettings` from a control callback. Only the successful commit pipeline advances baselines.

## Test the behavior, not only the draft

Add focused coverage that proves:

- the visible control exposes the required UI Automation pattern and a unique accessible name;
- real UIA/keyboard/pointer interaction changes the intended draft value;
- the dialog becomes dirty and the visible Apply button becomes enabled;
- restoring the baseline disables Apply;
- Apply updates the running settings snapshot and a fresh `TryLoadSettingsNoRecovery(...)` sees the serialized value;
- unrelated settings survive the merge;
- Apply advances the baseline and disables Apply;
- Cancel discards unapplied changes;
- settings round-trip and default pruning work;
- schema and localization contracts remain green.

Use existing shared UIA, settings-artifact backup, focus-wait, and TestSandbox helpers. Do not create page-local copies.

## Validate

Run, at minimum:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
Invoke-Pester -Script .\Tools\Tests\ResourceLocalizationContracts.Tests.ps1
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
.\.build\x64\Debug\SettingsSchemaTests.exe
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=<focused-case>
```

Run broader suites when the page changes shared shell, scrolling, theming, focus, or commit behavior.
