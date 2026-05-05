---
name: file-actions
description: File-action ID rules for viewer/editor/user-menu settings, parameterized commands, resolver behavior, and regression coverage. Use when changing file-action settings, direct dispatch, Edit New routing, or related specs/tests.
---

# file-actions

## Core contract

- File-action IDs preserve their configured casing when saved and displayed.
- All file-action ID comparisons are case-insensitive.
- IDs must be unique case-insensitively within each action list.
- IDs that differ only by case are invalid duplicates and must be rejected by settings validation.
- Association references and parameterized command IDs (`cmd/pane/viewWith/<id>`, `cmd/pane/editWith/<id>`, `cmd/pane/userMenu/<id>`) resolve case-insensitively.
- When lookup succeeds, downstream UI/runtime behavior should use the matched action definition rather than the caller-supplied casing.

## Change checklist

- Update settings-store validation and lookup together.
- Update runtime resolver and direct-dispatch call sites together.
- Keep preferences/action-status helpers aligned with the same case-insensitive lookup rule.
- Add focused selftests for:
  - persisted round-trip or no-recovery load success when references differ only by case
  - malformed-settings rejection for duplicate IDs that differ only by case
  - parameterized command dispatch with mixed-case configured IDs
- Update the authoritative docs:
  - `Specs/Core/Core_SettingsStore.md`
  - `Specs/SettingsStore.schema.json`
  - `Specs/UI/UI_CommandMenuKeyboard.md`
  - `Specs/Testing/Testing_TestCoverage.md`
