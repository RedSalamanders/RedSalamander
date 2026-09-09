# Advisor Plan 027 - ViewerImgRaw correctness bundle

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/027-viewerimgraw-correctness-bundle.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/027-viewerimgraw-correctness-bundle.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 027: ViewerImgRaw — menu ownership, WIC orientation, and TIFF scalar validation

## Status

- **Priority**: P2 (the double-DestroyMenu) / P3 (the rest)
- **State**: DONE — implementation, focused scheduler/terminal proof, consolidated archive, and repository-wide Full gate green (2026-07-12)
- **Effort**: S (each)
- **Risk**: LOW
- **Depends on**: none
- **Category**: bug / security (defense-in-depth)
- **Planned at**: commit `34c63e1d5`, 2026-06-21 (Decode.cpp anchors for items 2–3 re-refreshed in
  reconcile round 4 after plan 017's dispatch-hardening also added lines to this file; findings intact
  — originally `b274022d9`, 2026-06-16, refreshed once at `c3e89e580` for plan 015's CoInit block)

## Why this matters

Three narrow ViewerImgRaw issues that the corrected audit confirmed:
1. **Double `DestroyMenu` on menu-host attach failure** (correctness, P2): on a failure path the HMENU
   gets two owners and is destroyed twice → handle/heap corruption.
2. **EXIF orientation ignored for all WIC-decoded formats** (TIFF/PNG/HEIF-via-WIC): rotated images
   display the wrong way up.
3. **TIFF scalar IFD tags accept malformed declared types/counts**; the old `count*size` selection can
   wrap (defense-in-depth on untrusted input).

The former fourth item, **same-path duplicate decode**, is a possible optimization rather than a safe
point fix. Prefetch is request-generation-bound and does not deliver a UI completion; simply skipping
the main decode when the set contains a path can strand the viewer in loading state. It is retired
unless a future design shares a tokenized result and proves completion/cleanup on every exit path.

The final closeout audit supplied that missing design requirement and found a separate terminal-result
bug: every navigation launched another main decode, while a successful embedded thumbnail in
Thumbnail mode posted a non-final preview and then returned with the default final `E_FAIL`. The live
extension now uses one active plus one replaceable latest pending request, and Thumbnail mode moves a
successful embedded/sidecar thumbnail directly into one final success. RAW mode retains non-final
preview behavior.

### 2026-07-12 proof reconciliation

`TestViewerImgRawLatestWinsExactReaderAndCloseSafety` and
`TestViewerImgRawEmbeddedThumbnailTerminalSequencing` now pass. They prove pending replacement-count
telemetry, exact-reader rejection, current allocation/post terminal fallback, loading/in-flight clear,
blocked close, callback-return module lifetime, explicit shutdown quiet-point unload, and explicit
preview/final ordinals. The original two-source-file limit remains obsolete: shared Debug contracts,
tests, resources/localization, source contracts, performance instrumentation, and the authoritative
ViewerImgRaw spec are required proof. Consolidated proof is archived at
`Specs/TestRuns/4cb089111a23/Viewers/2026-07-12_105903_farsight_viewer_closeout/`: all 13 focused
cases passed with zero conditional skips and 1,426 performance metric records, including both new
ViewerImgRaw scheduler/terminal cases.

## Original state (superseded by the implementation below)

**(1) Double-DestroyMenu** — the window owns the menu *and* `_menuHandle` owns it; only a successful
`Attach` (which calls `SetMenu(hwnd, nullptr)`) reconciles to one owner, but its result is discarded:
```cpp
// ViewerImgRaw.cpp:3699,3707  (Open: window takes ownership of the HMENU)
HWND window = CreateWindowExW(..., menu.get(), g_hInstance, this);
if (! embeddedMode) { menu.release(); }              // window owns the HMENU
```
```cpp
// ViewerImgRaw.cpp:1476,1504  (OnCreate: _menuHandle ALSO owns it; Attach result ignored)
if (! _menuHandle) { _menuHandle.reset(GetMenu(hwnd)); }     // second owner (DestroyMenu)
// ...
static_cast<void>(_menuBarHost.Attach(g_hInstance, hwnd, _menuHandle.get()));   // SetMenu(null) only on success
```
On teardown, if `Attach` failed, the OS destroys the window's menu and `_menuHandle.reset()` (`:1337-1338`)
calls `DestroyMenu` again on the same handle.

**(2) WIC EXIF orientation** — the non-JPEG WIC branch hardcodes orientation = 1:
```cpp
// ViewerImgRaw.Decode.cpp:2655-2656  (after a successful WIC decode)
result->exif.orientation = 1;
result->exif.valid       = false;
```
`DecodeImageToBgraWic` (`:664-785`) does not read the WIC frame's orientation metadata. (Plan 015
added a CoInit block at the top of this function `:679-691`; it did NOT add any orientation read, so
this finding stands. Plan 017 added null-checks/Release to the dispatchers far below this function,
not to it.)

**(3) Retired duplicate-decode observation** — the main open worker inserts into `_inflightDecodes` unconditionally
(`~:2434-2437`, insert at `:2436`) without first checking whether a prefetch worker is already
decoding that path (the prefetch path does check `_inflightDecodes`, `~:2063-2069`: find at `:2064`,
insert at `:2068`). A set-only skip is not safe because prefetch has no UI completion contract.

**(4) TIFF scalar count/type validation** — bounds checks multiply a 32-bit attacker-controlled count without overflow
guard:
```cpp
// ViewerImgRaw.Decode.cpp:149  (ReadTiffShort)   if (count*2u <= 4u) ...
// ViewerImgRaw.Decode.cpp:178  (ReadTiffLong)    if (count*4u <= 4u) ...
```
(The downstream reads are individually `InRange`-guarded, so this is defense-in-depth, not a live
memory-safety hole — but it should not silently accept malformed counts.)

## Commands you will need

| Purpose | Command | Expected |
|---|---|---|
| Build plugin | `.\build.ps1 -ProjectName ViewerImgRaw` | exit 0, 0 warnings/errors |
| Full build | `.\build.ps1` | exit 0 |
| Tests | `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` | all pass |

## Scope

**In scope**: `Plugins/ViewerImgRaw/ViewerImgRaw.cpp`, `Plugins/ViewerImgRaw/ViewerImgRaw.Decode.cpp`,
focused test-only diagnostics in `Common/WindowMessages.h`, ViewerPE regression coverage, and the
authoritative ViewerImgRaw/Farsight documentation. **Out of scope**: the export UAF (014), WIC CoInit
(015), and unrelated dispatch hardening (017). The closeout extension explicitly includes the now-
defined scheduler/result-delivery contract; a set-only duplicate-decode skip remains forbidden.

## Steps

### Step 1: Fix the double-DestroyMenu (P2)

Call `NativeMenuBarHost::Attach` while the menu remains window-owned. Adopt the handle into
`_menuHandle` only when `Attach` succeeds **and** `GetMenu(hwnd) == nullptr` proves that the host
detached it. On failure, detach any partial DxUi host and leave the menu visible and solely
window-owned. A test-only fault seam must exercise repeated fallback open/close cycles.

**Verify**: `.\build.ps1 -ProjectName ViewerImgRaw` → exit 0.

### Step 2: Apply WIC frame orientation

In `DecodeImageToBgraWic`, read the frame's orientation (WIC `System.Photo.Orientation` via
`IWICMetadataQueryReader`, or apply it during decode) and propagate it so the non-JPEG branch sets the
real orientation instead of hardcoded 1. If reading WIC metadata is involved, an acceptable simpler fix
is to apply the orientation transform during the WIC decode pipeline (rotate the bitmap) so the produced
BGRA is already upright; then `orientation = 1` is honest. Pick whichever fits the existing decode flow.

**Verify**: open a rotated TIFF/PNG-with-orientation and confirm correct display (manual or test asset).

### Step 3: Retire the unsafe duplicate-decode point fix

Do not add a set-only check that skips the main decode. A future coalescing change must use a
generation-aware shared result and prove exactly one UI completion plus deterministic set/result
cleanup for success, stale request, post failure, and teardown.

**Verify**: source inspection confirms no new skip was added to the main-open path.

### Step 3b: Implement bounded latest-wins main scheduling and terminal delivery

Use one threadpool callback per viewer scheduler, with one active request and one replaceable pending
request. Each request owns the viewer/filesystem captures needed to outlive `Close`; pending
replacement and teardown release those owners without waiting. Check generation before every whole-
file read and expensive decoder entry. Allocation or final payload-post failure records an
allocation-free shared terminal slot polled by the loading timer, clears the current in-flight marker,
and never dereferences a destroyed/recycled `HWND`. Threadpool module pins transfer through
`FreeLibraryWhenCallbackReturns`.

**Verify**: a blocked provider observes one active main read and one coalesced latest pending request;
close remains bounded; forced result-allocation and post failures both leave loading false and
in-flight count zero; queue/resource metrics prove the one-active invariant.

### Step 3c: Make successful Thumbnail mode terminal exactly once

When a sidecar or embedded thumbnail succeeds and Thumbnail mode is requested, publish one final
success containing that thumbnail. Keep the non-final preview only when RAW mode will continue to a
full decode.

**Verify**: the deterministic embedded-thumbnail seam/fixture observes one final success, no warning
alert, and no trailing default `E_FAIL`; RAW mode still observes preview then final RAW.

### Step 4: Overflow-safe TIFF count checks

For scalar SHORT/LONG/RATIONAL tags, require the expected TIFF type and `count == 1` before reading.
Read the inline scalar value directly (or the rational's checked offset) so no attacker-controlled
`count * N` multiplication selects the storage path.

**Verify**: `.\build.ps1` → exit 0.

## Test plan

- `TestViewerImgRawMenuOwnershipOrientationAndExifScalarGuards` exercises three forced native-menu
  fallback cycles, WIC TIFF orientation, RAW-extension→WIC fallback orientation, valid scalar EXIF
  orientation, and wrong-type/huge-count rejection.
- Existing `TestViewerImgRawDecodesPngThroughWicWithoutErrorAlert` and
  `TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo` remain green.
- `TestViewerImgRawLatestWinsExactReaderAndCloseSafety` provides latest-wins/blocked-provider coverage,
  exact-reader hostile cases, forced
  `REDSALAMANDER_VIEWERIMGRAW_FORCE_OPEN_RESULT_ALLOCATION_FAILURE` and
  `REDSALAMANDER_VIEWERIMGRAW_FORCE_OPEN_RESULT_POST_FAILURE` coverage, callback-return/module-unload
  proof, and scheduler replacement/resource telemetry.
- `TestViewerImgRawEmbeddedThumbnailTerminalSequencing` proves explicit preview/final ordinals,
  Thumbnail final-success-without-alert, and RAW preview-then-final behavior.

## Done criteria

- [x] On `Attach` failure, the HMENU has exactly one owner and the native fallback remains usable.
- [x] WIC-decoded images honor EXIF/WIC orientation, including RAW→WIC fallback.
- [x] The unsafe set-only duplicate-decode proposal is explicitly retired.
- [x] Scalar TIFF reads validate type/count without attacker-controlled multiplication.
- [x] `.\build.ps1 -ProjectName ViewerImgRaw` and `.\build.ps1 -ProjectName ViewerPETests` exit 0 with 0 warnings.
- [x] Focused and existing ViewerImgRaw regressions pass.
- [x] `plans/README.md` reflects the reconciled implementation.
- [x] Main opens are limited to one active plus one replaceable latest pending request; close/cancel never waits.
- [x] Current result allocation/post failure has an allocation-free UI fallback and clears loading/in-flight state.
- [x] Successful Thumbnail mode produces one final success; RAW mode preserves preview semantics.
- [x] Focused scheduler/fault/thumbnail/blocked-close/callback-unload proof and scoped builds pass.
- [x] Consolidated Farsight metrics/archive evidence is recorded.
- [x] `.\Tools\Run-AllTests.ps1 -Suite Full` passes in the final closeout workspace (`Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`).

## STOP conditions

- Excerpts don't match live code (drift).
- `NativeMenuBarHost::Attach`'s success path does not actually `SetMenu(null)` — re-derive the correct
  single-owner reconciliation and report if unclear.
- Reading WIC orientation metadata requires a large refactor — land Steps 1, 3, 4 and defer Step 2 with
  a note.

## Maintenance notes

- Item 1 reviewer focus: HMENU has exactly one RAII owner on every Open/teardown path.
- Items 2–4 are quality/defense-in-depth; none changes the (already-sound) untrusted-decode safety
  posture established by the decode bounds checks and dimension caps.
