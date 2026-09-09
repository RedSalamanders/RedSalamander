# Advisor Plan 022 - ViewerVLC full-path libVLC load

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/022-viewervlc-dll-search-path.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/022-viewervlc-dll-search-path.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 022: ViewerVLC — full-path libVLC loading without process-global DLL-search mutation

## Status

- **State**: DONE — implementation, focused proof, installed-VLC proof, consolidated archive, and repository-wide Full gate green (2026-07-12)
- **Priority**: P2 (security and cross-thread process-state race)
- **Effort**: M
- **Risk**: MED (libVLC dependency and cleanup lifetime)
- **Category**: security / concurrency
- **Planned at**: commit `b274022d9`, 2026-06-16

## Historical defect

The load worker saved and changed the process-global, non-additive DLL directory for the duration of a
playback state. Two panes could race that state, and unrelated host/plugin DLL loads could resolve from
VLC's folder. `libvlc.dll` was already addressed by full path; only sibling dependency resolution
needed a scoped solution.

## 2026-07-12 implementation reconciliation

Initial load now uses the full `libvlc.dll` path with:

```cpp
LoadLibraryExW(dllPath.c_str(), nullptr,
               LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS)
```

The implementation removes `GetDllDirectoryW`, `SetDllDirectoryW`, restore state, and the old
`previousDllDirectory`/`dllDirectoryWasSet` fields. It does not use the additive-directory fallback.
The complete `VlcState` moves through async cleanup, and its `wil::unique_hmodule` remains alive until
player, media, and instance cleanup finishes.

The final security pass expanded discovery and configuration boundaries adjacent to the original
finding:

- automatic discovery is limited to registered VLC paths and
  `FOLDERID_ProgramFiles` / `FOLDERID_ProgramFilesX86`;
- `SearchPathW` and process-environment Program Files lookup are forbidden;
- enumeration-like vout/aout values are bounded tokens;
- one leading `no-` is normalized case-insensitively before deny matching;
- logfile/file-logging/pidfile plus plugin/search/config/interface/module/codec/access/filter/output
  overrides are rejected;
- signed and unsigned JSON numbers are clamped before narrowing.

The persistent cleanup architecture is specified by corrected plan 023.

## Reconciled scope

The durable change spans `ViewerVLC.cpp`/`.h`, configuration resources/schema, shared Debug contracts,
ViewerPETests, source/localization contracts, and the authoritative ViewerVLC spec. The original
“only `ViewerVLC.cpp`/`.h` modified” criterion is obsolete as of 2026-07-12.

## Requirements

1. Load `libvlc.dll` by full path with `LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR` and default directories.
2. Keep the loaded module in `VlcState` through complete player/media/instance teardown.
3. Never mutate or query the process DLL directory from production ViewerVLC code.
4. Never use global path search or environment-variable Program Files discovery.
5. Reject configuration tokens that can redirect dependency/plugin/module search or external output.
6. Keep all discovery, loading, export resolution, and `libvlc_new` work off the UI thread.

## Focused proof

- Source contracts enforce one full-path load site and ban global DLL-directory APIs plus
  `SearchPathW`.
- Configuration boundary tests cover mixed-case negation, logfile/file-logging/pidfile rejection,
  separator/option injection, and maximum-`uint64` clamping.
- Scoped ViewerVLC builds are zero-warning; focused configuration tests are green.
- The installed-VLC path verified a real module/instance/player, unchanged process DLL-directory
  state, Tab focus transfer to the HUD, and clean close.
- Consolidated proof is archived at
  `Specs/TestRuns/4cb089111a23/Viewers/2026-07-12_105903_farsight_viewer_closeout/`: all 13 focused
  cases passed with zero conditional skips, including all three ViewerVLC cases, and the archive
  contains 1,426 performance metric records.

## Done criteria

- [x] No production `GetDllDirectoryW`/`SetDllDirectoryW`/`AddDllDirectory`/`RemoveDllDirectory` or
      `SearchPathW` call remains in ViewerVLC.
- [x] libVLC loads by full path with per-module dependency search.
- [x] `VlcState` owns the module through player/media/instance cleanup.
- [x] Registry/known-folder discovery replaces environment/global search.
- [x] Dangerous configuration and negated dangerous options are rejected.
- [x] Scoped zero-warning build, source contracts, and focused configuration regressions pass.
- [x] Expanded tests/resources/spec scope is recorded; the old two-file-only criterion is retired.
- [x] `plans/README.md` reflects this reconciliation.
- [x] Refreshed installed-VLC playback/Tab/HUD proof is green where VLC is present.
- [x] Consolidated VLC/Farsight metrics archive is recorded.
- [x] `.\Tools\Run-AllTests.ps1 -Suite Full` passes in the final closeout workspace (`Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`).

## STOP conditions

- A dependency can load only by restoring process-global DLL search mutation.
- Cleanup would release the VLC module before player/media/instance teardown.
- Any load/discovery fallback runs synchronously on the UI thread.

## Maintenance notes

The plugin rule is stronger than merely banning `SetDllDirectoryW`: resolve the trusted top-level
module by exact path, keep dependency resolution scoped to that load, and keep automatic discovery to
trusted registry/known-folder roots.
