# Advisor Plan 029 - Viewer hardcoded-string localization

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/029-viewer-hardcoded-strings-localization.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/029-viewer-hardcoded-strings-localization.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 029: Viewer plugins — resource-back all user-facing narration and controls

## Status

- **State**: DONE — implementation, focused localization proof, consolidated archive, and repository-wide Full gate green (2026-07-12)
- **Priority**: P3 (localization correctness and consistency)
- **Effort**: M
- **Risk**: LOW
- **Category**: localization / maintainability
- **Planned at**: commit `b274022d9`, 2026-06-16

## Historical and expanded findings

The original audit found hardcoded English on ViewerImgRaw export/decode paths and ViewerSqlite error
paths. Farsight closeout extended the inventory:

- **ViewerImgRaw**: exact whole-file reader failures, LibRaw policy/failure narration, async terminal
  delivery, WIC export failures, and export success;
- **ViewerSqlite**: SQLite failure context, statement guards, snapshot/limit/cancellation diagnostics;
- **ViewerPE**: rendered and exported report headings, labels, and error narration;
- **ViewerWeb**: generated JSONL summaries, state words, controls, and error narration.

These strings reach UI status, alerts, reports, or generated documents. Debug-only diagnostic labels
remain source literals. Canonical format identifiers (for example PE/SQLite structural tokens) may
remain canonical data when they are not narrative UI text.

## 2026-07-12 implementation reconciliation

User-facing text is defined by `IDS_*` entries and loaded with `LoadStringResource` or
`FormatStringResource`. Formatted resources use indexed positional placeholders (`{0}`, `{1:08X}`)
with exact token parity across base and cs-CZ/fr-FR/ja-JP/sk-SK satellites. End-user ImgRaw messages
drop the internal `ViewerImgRaw:` prefix while Debug logs retain component context.

ViewerWeb places localized generated-document strings into escaped data and assigns them through
`textContent`; localization does not reopen HTML/script injection. ViewerPE report generation keeps
canonical values separate from localized narration so export and on-screen reports match.

## Reconciled scope

The implementation spans the four plugins' source, `resource.h`, base/satellite `.rc` files, project
resource wiring where required, ViewerPETests/runtime assertions, source/localization contracts, and
authoritative plugin specs. The original two-plugin scope and implicit “resource files only” criterion
are obsolete as of 2026-07-12.

## Requirements

1. Inventory every string that reaches alerts, status, rendered/exported reports, or generated viewer
   controls/summaries; exclude Debug-only logs.
2. Add stable `IDS_*` entries and load/format through the project localization helpers.
3. Use indexed positional placeholders in source-string argument order.
4. Preserve exact placeholder tokens in every satellite; translations may reorder but may not add,
   remove, duplicate, renumber, or change format specs.
5. Keep localized Web data escaped and assigned as text, never interpreted as markup/script.
6. Keep canonical file-format tokens distinct from narrative localization.

## Focused proof

- The focused source-contract lane passes 136/136 assertions across the Farsight viewer contracts.
- The focused localization lane passes 4/4, including placeholder parity/resource ownership cases.
- Compiled localization tests and scoped ViewerImgRaw/ViewerSqlite/ViewerPE/ViewerWeb builds are green.
- Focused viewer failure/report paths resolve non-empty resource-backed text.
- Consolidated proof is archived at
  `Specs/TestRuns/4cb089111a23/Viewers/2026-07-12_105903_farsight_viewer_closeout/`: all 13 focused
  cases passed with zero conditional skips and 1,426 performance metric records.

## Done criteria

- [x] ViewerImgRaw reader/LibRaw/terminal/export user text is resource-backed.
- [x] ViewerSqlite user-facing failure, guard, snapshot, and budget text is resource-backed.
- [x] ViewerPE rendered/exported report narration is resource-backed.
- [x] ViewerWeb generated JSONL narration/controls are resource-backed and injected as escaped text.
- [x] Base and cs/fr/ja/sk resources preserve exact indexed-placeholder parity.
- [x] Debug-only logs remain diagnostic literals and canonical values remain distinguishable.
- [x] Focused source contracts, localization contracts, compiled localization, and scoped builds pass.
- [x] Expanded four-plugin/test/spec scope is recorded; the original scope criterion is retired.
- [x] `plans/README.md` reflects this reconciliation.
- [x] Operation Farsight's consolidated closeout evidence is finalized.
- [x] `.\Tools\Run-AllTests.ps1 -Suite Full` passes in the final closeout workspace (`Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`).

## STOP conditions

- A resource uses bare `{}` or an unindexed format spec.
- A satellite changes the source placeholder token set.
- Localized ViewerWeb text is inserted through `innerHTML` or an executable script literal.
- A canonical structural token is translated as though it were narrative UI text.

## Maintenance notes

“No hardcoded UI strings” applies equally to generated reports/documents and ordinary Win32 controls.
New viewer behavior should add the resource and satellite placeholder contract in the same change.
