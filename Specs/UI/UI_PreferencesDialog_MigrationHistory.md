# Preferences Dialog Migration History

> **Status:** RETIRED historical record
> **Retired:** 2026-08-04
> **Baseline:** `fc4ba4a63`

The phased Preferences migration is complete and this path is retained because
older specifications and historical plans link to it. It is not a current
contract, implementation queue, validation checklist, or source of product
requirements.

Current authority:

- `Specs/UI/UI_PreferencesDialog.md` — Preferences behavior and integration;
- `Specs/UI/UI_DxUiSharedGrid.md` — shared DxUi controls, accessibility, and
  migrated-window acceptance;
- `Specs/Core/Core_SettingsStore.md` and
  `Specs/SettingsStore.schema.json` — settings persistence and schema;
- `Specs/Plans/WIP/Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md`
  — unresolved optional product choices recovered during retirement.

Retrieve the full pre-retirement record with:

```powershell
git show fc4ba4a63:Specs/UI/UI_PreferencesDialog_MigrationHistory.md
```

Implemented behavior was reconciled into the current authority above. Completed
bug diaries, phased task lists, resume notes, and obsolete refactor proposals
remain available only through Git history.
