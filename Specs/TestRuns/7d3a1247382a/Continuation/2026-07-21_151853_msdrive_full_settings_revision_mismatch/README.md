# Microsoft Drive closeout Full-run continuation

## Context

- Date: 2026-07-21
- Branch: `codex/msdrive-mutation-reconciliation`
- Base implementation commit: `5242378b7`
- Command: `.\Tools\Run-AllTests.ps1 -Suite Full`
- TestSandbox run ID: `20260721T122429Z-19172-3aadd909e4dd4b1ca03c9e9035cb9909`
- Purpose: broad closeout gate for the Microsoft Drive ambiguous PATCH/DELETE reconciliation optimization.

The user requested a pause while the Full run was still executing. The active wrapper and its File Operations child
were stopped only after their current artifacts had been copied here. This is continuation evidence, not a passing
Full result.

## Preserved result

Commands completed with 526 passed, 18 failed, and 0 skipped. Two test-owned modal dialogs blocked the run until
dismissed after evidence capture:

- `RedSalamander-debug.settings.json`: `Could not open the settings file`, `0x8007051A`.
- `RedSalamanderMonitor-debug.settings.json`: the same error and HRESULT.

`0x8007051A` is `HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH)`. The earliest structured Commands failure is
`settings_store_search_roundtrip`: after a successful save, the test constructs a fresh unstamped `Settings` snapshot
and attempts to overwrite the same app-ID target. The SettingsStore CAS correctly rejects that stale/missing expected
stamp. A focused continuation must repair the test to carry the committed stamp by mutating/reusing the saved snapshot,
then rerun Commands to determine which later settings failures remain independent.

The completed Commands result also contains three visibly independent failures that must not be attributed to the
Microsoft Drive provider change: Quick Search focus, isolated SearchService readiness, and a drop-integrity capability
expectation. See `commands-results.json` and `commands-trace.txt` for the exact cases.

File Operations had reached 3 passed / 0 failed in its current partial result when the pause was requested. Its result
and trace are intentionally named `fileops-partial-*`; they are not an acceptance gate.

## Continuation order

1. Fix `settings_store_search_roundtrip` so its second CAS-protected save reuses the mutable snapshot whose committed
   stamp was advanced by the first save.
2. Run that exact Commands case, then the settings/Connection Manager/Preferences affected slice.
3. If the settings-file link dialogs remain, inspect test artifact backup/restore and coordinator lineage around
   missing-file restoration; do not weaken production CAS.
4. Rerun the Microsoft Drive focused Debug/test-enabled Release contracts only if code changes overlap the provider.
5. Complete a trustworthy broad gate or classify every remaining unrelated failure with durable evidence before
   moving the WIP plan to Done.
