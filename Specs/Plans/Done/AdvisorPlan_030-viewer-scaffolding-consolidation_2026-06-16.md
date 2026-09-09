# Advisor Plan 030 - Viewer scaffolding consolidation

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/030-viewer-scaffolding-consolidation.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/030-viewer-scaffolding-consolidation.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 030: Consolidate duplicated viewer scaffolding (D2D/theme/window/async) into Common — staged refactor

## Status

- **State**: DONE — Stage 1/2 implementation, focused quiet-point proof, consolidated archive, final DxUi execution, and repository-wide Full gate green (2026-07-12)
- **2026-07-12 reconciliation**: concrete Stage 1/2 defects are implemented; Stages 3–5 are retired from Operation Farsight because the live async/window/rendering models have diverged and those stages are optional architecture work, not correctness remediation. See the closeout note below.
- **Priority**: P3 (concrete correctness/RAII items only; broader dedup moved to backlog)
- **Effort**: L (whole plan) — staged S→M per item
- **Risk**: MED (touches all 7 viewer plugins; mechanical but wide blast radius)
- **Depends on**: pairs with existing plan 005 (Factory.cpp dedup); do 005 first or alongside
- **Category**: tech-debt / architecture
- **Planned at**: commit `b274022d9`, 2026-06-16

## 2026-07-12 live-code reconciliation

- **Stage 1 implemented:** `Common/DxUi` already contained the one alpha-preserving `ColorFromArgb` implementation, but it was private to `DxUi.Internal.h`. It is now a public `DxUi.h` primitive, ViewerPE routes themed Direct2D clear/background/text colors through it, and `DxUiTests` pins alpha plus channel order. This fixes the only observable drift (ViewerPE dropping ARGB alpha) without duplicating another converter.
- **Stage 2 implemented to the live defect:** ViewerText, ViewerImgRaw, and ViewerWeb had already migrated their class-background state to static `wil::unique_hbrush` owners. ViewerSpace was the only remaining raw heap owner, and its state is now a static RAII object with explicit shutdown reset. The closeout unload audit additionally required ViewerImgRaw to unregister its window class and reset both active/pending brushes at `RedSalamanderPluginShutdown`; `RedSalamanderPluginCanUnloadNow` now gates unload on that quiet point plus zero queued callbacks. No brush cleanup runs from `DllMain`.
- **Stage 3 retired:** the point fixes now use materially different schedulers and ownership contracts (latest-wins detached PE scheduling, payload-owned Web/Text results, ImgRaw decode/export reservations). A single `std::function` helper would erase module-pin, cancellation, terminal-delivery, and resource-budget distinctions rather than make them safer.
- **Stage 4 retired from this remediation:** it was explicitly maintainer-gated design-spike work and has no unresolved behavior defect after Stage 1. It belongs in an architecture backlog with an approved abstraction proposal, not in a safety closeout.
- **Stage 5 retired:** it was optional mechanical dedup with no defect. The current per-viewer navigation state remains covered by the cross-viewer harness.

Operation Farsight treats the source architecture as complete. Focused source/quiet-point contracts,
ImgRaw callback-unload proof, and the consolidated 13/13 focused archive are green. Retired stages are
not silently deferred requirements.

## Why this matters

The seven viewer plugins independently re-implement the same scaffolding. None of this is a bug (the
audit verified the copies behave correctly and the "god-file" line counts were overstated), but the
duplication is real maintenance drag: a fix or theme/ABI change must be applied in up to seven places,
and the copies drift (e.g. ViewerPE's color conversion drops alpha). Consolidating the **mechanical,
low-risk** pieces into `Common/` removes that drag and makes the next ABI/theme change a one-place edit.

Confirmed duplication (from the deep audit):
- **ARGB→D2D color conversion** repeated per plugin; **ViewerPE's variant drops alpha**
  (`ViewerPE.cpp:251-257`, used `:1660`).
- **Global background-brush singleton** using raw `new`/`delete` + global mutable state, copied in
  ViewerSpace (`:321-347`), ViewerText (`:4053-4100`), ViewerImgRaw (`:466-511`), ViewerWeb (`:641-685`).
  (The audit confirmed the cross-thread access is UI-thread-affine, so this is style/dedup, not a race —
  but the raw `new`/`delete` violates the RAII rule.)
- **`RegisterWndClass` (static ATOM) + `WndProcThunk` (WM_NCCREATE stash + InitPostedPayloadWindow)**
  boilerplate copied across all 7 plugins (e.g. `ViewerPE.cpp:650-690`).
- **D2D factory + HWND render target + device-lost (`D2DERR_RECREATE_TARGET`) setup/teardown** repeated
  per plugin (ViewerPE/Text/Space/ImgRaw all link `Common/DxUi`; ViewerVLC uses a different model).
- **`otherFiles` navigation** (`_otherFiles`/`_otherIndex` + Next/Prev/First/Last) re-implemented in
  ViewerText, ViewerPE, ViewerWeb, ViewerImgRaw.
- **Async work-item dispatch** (AddRef + module-keepalive + `TrySubmitThreadpoolCallback` +
  `PostMessagePayload` + `requestId`) — see plan 017; a shared helper would make the AddRef/Release
  balance impossible to get wrong.

## Commands you will need

| Purpose | Command | Expected |
|---|---|---|
| Build all viewers | `.\build.ps1` | exit 0, 0 warnings/errors |
| Per plugin | `.\build.ps1 -ProjectName Viewer<Name>` | exit 0 |
| Tests | `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` | all pass |

## Scope

**Reconciled in scope**: expose the existing `DxUi::ColorFromArgb` helper publicly; route ViewerPE
theme conversion through it; replace ViewerSpace's remaining raw class-brush state with WIL RAII and
an explicit shutdown reset; and require ViewerImgRaw class/brush cleanup only at its explicit shutdown
quiet point with an honest zero-callback unload vote. Required tests, source contracts, shared headers,
and authoritative specs are part of this scope.

**Retired/out of scope**: a generic async dispatcher, a shared window/D2D surface rollout, ViewerVLC's
different HUD/overlay render model, and mechanical `otherFiles` navigator dedup. Those require an
independent approved architecture proposal and are not Farsight done criteria.

## Git workflow

- Branch per stage: `advisor/030-stageN-<slug>`
- Message: `refactor(viewers): hoist <thing> into Common (no behavior change)`
- Do NOT push/PR unless instructed.

## Stages (each independently shippable)

### Stage 1 (implemented): public ARGB→D2D color helper + PE alpha fix

`DxUi::ColorFromArgb` was made public in `Common/DxUi.h`, and ViewerPE now routes its themed Direct2D
colors through that existing alpha-preserving implementation. `DxUiTests` pins alpha and channel order.

**Verify**: `.\build.ps1` → exit 0; viewer theme cases in `Tests/ViewerPETests` green.

### Stage 2 (implemented to the live defects): RAII class brushes and explicit quiet-point reset

The live survey showed ViewerText, ViewerImgRaw, and ViewerWeb already used static
`wil::unique_hbrush` owners, so forcing another shared wrapper would add churn without removing a
defect. Replace ViewerSpace's remaining raw state with static WIL RAII and explicit plugin-shutdown
reset. ViewerImgRaw must unregister its class and reset active/pending brushes only at
`RedSalamanderPluginShutdown`; `CanUnloadNow` requires that quiet point plus zero callbacks. `DllMain`
performs no cleanup.

**Verify**: `.\build.ps1` → exit 0; viewers render correct backgrounds.

### Stage 3 (retired): shared async-dispatch helper

Do not implement the old `SubmitViewerWork(IUnknown*, std::function<...>)` proposal. PE uses detached
latest-wins scheduler state; ImgRaw has distinct main/prefetch/export reservations; Text/Web have
different payload, terminal, and budget contracts; VLC owns a persistent cleanup dispatcher. A generic
wrapper would hide the very ownership distinctions the closeout made explicit.

### Stage 4 (retired from Farsight): shared window-class/thunk + D2D HWND surface

The old shared `RegisterWndClass`/`WndProcThunk` and `HwndD2DSurface` rollout is not a Farsight
requirement. Reconsider only as a separately approved design spike with independent maintenance value;
ViewerVLC remains excluded because its retained video/HUD/overlay model is materially different.

### Stage 5 (retired optional work): shared `otherFiles` navigator

No behavior defect justified this mechanical churn. Reconsider only in an independent refactor.

## Test plan

- `DxUiTests` contains the ARGB alpha/channel-order regression and must run in the final lane.
- Focused ViewerPE/ViewerSpace/ViewerImgRaw and source-contract tests cover the affected call sites,
  RAII brush state, explicit shutdown, callback quiet point, and clean unload.
- Scoped builds are zero-warning.

## Done criteria (per stage)

- [x] The existing `Common/DxUi::ColorFromArgb` implementation is public and ViewerPE uses it for themed Direct2D colors.
- [x] Focused source/quiet-point and ImgRaw callback-unload tests are green.
- [x] Stage 1 additionally: ViewerPE no longer drops alpha in Direct2D theme conversion.
- [x] Stage 2 additionally: no raw `new`/`delete` remains in the viewer class-background-brush code, and ViewerSpace/ViewerImgRaw reset class-brush state at explicit plugin shutdown quiet points.
- [x] Scoped ViewerPE/ViewerSpace/ViewerImgRaw builds are zero-warning.
- [x] `plans/README.md` reflects the reconciliation and retired stages.
- [x] Consolidated Farsight proof is archived at `Specs/TestRuns/4cb089111a23/Viewers/2026-07-12_105903_farsight_viewer_closeout/` (13/13 focused cases, zero conditional skips, 1,426 performance metric records).
- [x] Final `DxUiTests` alpha/channel-order execution passes in the Full gate.
- [x] Repository-wide `.\Tools\Run-AllTests.ps1 -Suite Full` passes at final closeout (`Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`).

## STOP conditions

- The cited duplication has materially changed since `b274022d9` (drift) — re-survey before extracting.
- A "shared" abstraction would force ViewerVLC (different render model) into an awkward shape — exclude
  VLC from that stage and note it.
- Any stage changes observable viewer behavior — revert and report; this plan must be behavior-preserving.

## Maintenance notes

- This plan is the natural home for the maintenance notes left by plans 017 (async helper), 021 (unified
  worker model), and 015 (per-thread COM helper). Land those point-fixes first; this plan generalizes
  them.
- The audit explicitly judged the "god-file" split and theme-version gating as **not** worth changing
  (correct as-is) — do not include them here.
- Reviewer: for each stage, diff the generated viewer behavior (theme colors, window class, async
  results) against pre-refactor; the bar is byte-identical behavior.
