# Advisor Plan 028 - ViewerText reliability and dead code

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/028-viewertext-reliability-and-deadcode.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/028-viewertext-reliability-and-deadcode.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 028: ViewerText — bounded hex clipboard and removal of unreachable/dead paths

## Status

- **State**: DONE — implementation, focused proof, consolidated archive, and repository-wide Full gate green (2026-07-12)
- **Priority**: P3 (UI responsiveness, allocation safety, maintainability)
- **Effort**: M after test/localization/performance expansion
- **Risk**: LOW–MED
- **Category**: reliability / tech debt
- **Planned at**: commit `34c63e1d5`, 2026-06-21

## Historical defects

1. A large hex selection could synchronously build an unbounded CSV `std::wstring` on the UI thread.
2. The optional “show as hex when text loading fails” block was unreachable because success was forced
   immediately before its `FAILED(hr)` guard.
3. `SaveCookie::StreamOutCallback`, `_msftEditModule`, and the duplicate
   `DetectEncodingAndSize` path were unused.

## 2026-07-12 implementation reconciliation

### Bounded clipboard planning

`ViewerTextSafety::ComputeHexClipboardPlan` uses checked `uint64_t` arithmetic and caps synchronous
work at `kMaxHexClipboardBytes == 256 KiB` of source data before the CSV string is reserved or
formatted. An exact-cap selection is accepted. A cap+1 or larger selection copies only the bounded
prefix, records rejected bytes, emits accepted/rejected byte and duration metrics, and shows exactly
one localized warning. Planning a `UINT64_MAX` selection remains constant-space and cannot wrap.

No `bad_alloc` catch was added. Allocation failure remains fatal under the repository exception
contract; the deterministic cap prevents attacker/user-controlled proportional allocation.

### Dead fallback decision

The unreachable automatic retry block and `allowHexFallback` plumbing were deleted. The authoritative
ViewerText contract now says an I/O/decode failure is terminal; successful binary-content detection may
still choose hex mode. This avoids turning a genuine provider/decode failure into a misleading second
read and removes inert control flow.

### Dead code removal

`SaveCookie::StreamOutCallback`, `_msftEditModule`, and `DetectEncodingAndSize` plus its declaration
were removed. Source contracts/grep pin their absence.

## Reconciled scope

The implementation spans ViewerText source/header files, `ViewerText.SafetyHelpers.h`, resource
strings/satellites, shared Debug requests/snapshots, ViewerPETests, metrics, source contracts, and the
authoritative ViewerText spec. The original “only three ViewerText files modified” criterion is
obsolete as of 2026-07-12 because localized user feedback and deterministic production-path proof are
part of the behavior contract.

## Focused proof

- `TestViewerTextDecodeAndClipboardSafetyHelpers` covers exact-cap, cap+1, and `UINT64_MAX` planning.
- The production debug-dispatch path copies an exact-cap selection without warning.
- The production over-cap path performs the same bounded source-byte work and emits one localized
  warning.
- Focused metrics on machine `4cb089111a23` record exact-cap 262,144 accepted / 0 rejected and cap+1
  262,144 accepted / 1 rejected.
- Grep/source contracts find no dead symbols, `allowHexFallback`, fallback diagnostic, or ViewerText
  `bad_alloc` catch.
- Scoped ViewerText/ViewerPETests builds and focused regressions are green.
- Consolidated proof is archived at
  `Specs/TestRuns/4cb089111a23/Viewers/2026-07-12_105903_farsight_viewer_closeout/`: all 13 focused
  cases passed with zero conditional skips and 1,426 performance metric records, including the
  ViewerText async/terminal and diff/clipboard coverage.

## Done criteria

- [x] Hex clipboard proportional work is capped at 256 KiB of source bytes.
- [x] Checked planning handles exact-cap, first-over-cap, and `UINT64_MAX` without wrap/allocation.
- [x] Truncation is visible through exactly one localized warning and accepted/rejected metrics.
- [x] No allocation-swallowing exception handler was introduced.
- [x] The unreachable automatic text-failure→hex block and plumbing are removed.
- [x] Successful binary-content auto-selection remains intact and documented separately.
- [x] `StreamOutCallback`, `_msftEditModule`, and `DetectEncodingAndSize` are absent.
- [x] Scoped zero-warning builds, focused helper/runtime tests, and source contracts pass.
- [x] Expanded helper/resource/test/spec scope is recorded; the old narrow-file criterion is retired.
- [x] `plans/README.md` reflects this reconciliation.
- [x] Consolidated ViewerText/Farsight metrics archive is finalized.
- [x] `.\Tools\Run-AllTests.ps1 -Suite Full` passes in the final closeout workspace (`Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`).

## STOP conditions

- Any path reserves/formats proportional to an unbounded selection before applying the cap.
- A `bad_alloc` or broad exception handler is added to continue after allocation failure.
- Automatic hex retry is restored without an explicit product-contract change and terminal/exact-read
  proof.

## Maintenance notes

The cap is expressed in source bytes, not output characters, so output expansion remains predictably
bounded. If a future feature needs full multi-megabyte export, it should be a cancellable asynchronous
file operation rather than an unbounded clipboard build on the UI thread.
