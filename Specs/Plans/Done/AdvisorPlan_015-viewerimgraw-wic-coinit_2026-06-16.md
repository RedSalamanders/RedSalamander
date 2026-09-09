# Advisor Plan 015 - ViewerImgRaw WIC CoInitialize

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/015-viewerimgraw-wic-coinit.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/015-viewerimgraw-wic-coinit.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 015: ViewerImgRaw — initialize COM on decode threadpool workers (fix silent failure of PNG/GIF/BMP/TIFF/HEIC)

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. Run every verification
> command and confirm the expected result before continuing. If a STOP condition occurs,
> stop and report. Historical: status-row updates no longer apply in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat b274022d9..HEAD -- Plugins/ViewerImgRaw/ViewerImgRaw.Decode.cpp`
> If the file changed, compare the "Current state" excerpts to the live code first; on a
> mismatch treat it as a STOP condition.

## Status

- **Priority**: P1 (entire image-format family is silently broken)
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: bug
- **Planned at**: commit `b274022d9`, 2026-06-16
- **Reviewed/tightened**: 2026-06-21 (HEAD `bd9e5c2bc`) — adversarial multi-agent review confirmed
  the bug and the fix; no source drift; corrected the test (the original child-window assertion
  could not detect the bug) to host-alert observability + a WIC-verified PNG fixture; resolved
  the include question (no change needed); scope/done-criteria updated to include the test file.

## Why this matters

`DecodeImageToBgraWic` creates a WIC factory via `CoCreateInstance(..., CLSCTX_INPROC_SERVER, ...)`,
which requires COM to be initialized on the calling thread. It is invoked **only** from
threadpool worker lambdas (`TrySubmitThreadpoolCallback` with a null callback environment),
and those threads have no COM apartment. None of the decode workers call `CoInitializeEx`
(only the *export* worker does, at `Export.cpp:723`). As a result `CoCreateInstance` returns
`CO_E_NOTINITIALIZED` and **every WIC-backed format — PNG, GIF, BMP, non-RAW TIFF, HEIC/WebP
via WIC codecs — silently fails to open**, and the RAW-decode fallback for corrupt RAWs is
also dead. JPEG (turbojpeg) and libraw RAW happen to work because they never touch COM, which
masks the bug in casual testing. For an image viewer, "PNG won't open" is a top-severity
reliability defect.

## Current state

- `Plugins/ViewerImgRaw/ViewerImgRaw.Decode.cpp` — `DecodeImageToBgraWic` and its callers.

```cpp
// ViewerImgRaw.Decode.cpp:664-684  (free function)
HRESULT DecodeImageToBgraWic(const uint8_t* data, size_t sizeBytes, uint32_t& outWidth,
                             uint32_t& outHeight, std::vector<uint8_t>& outBgra) noexcept
{
    // ...
    wil::com_ptr<IWICImagingFactory> factory;
    HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.addressof()));
    if (FAILED(hr))
        hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.addressof()));
    if (FAILED(hr) || ! factory)
        return FAILED(hr) ? hr : E_FAIL;   // <-- on a threadpool thread, hr == CO_E_NOTINITIALIZED
    // ...
}
```

Call sites, all inside threadpool worker lambdas:
- `ViewerImgRaw.Decode.cpp:2619` — non-JPEG branch of the main open worker (`StartAsyncOpen`).
- `ViewerImgRaw.Decode.cpp:2751` — RAW-failure fallback for non-JPEG extensions.
- `ViewerImgRaw.Decode.cpp:2100` — neighbor prefetch worker (`StartPrefetchNeighbors`).

**The correct pattern is already in this codebase**, in the export worker:

```cpp
// ViewerImgRaw.Export.cpp:723-731  (export worker lambda — copy this)
const HRESULT coinitHr  = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
const bool shouldUninit = SUCCEEDED(coinitHr);
auto coUninit           = wil::scope_exit([&]
{
    if (shouldUninit) { CoUninitialize(); }
});
```

WIC is apartment-agnostic, so MTA (`COINIT_MULTITHREADED`) is correct and matches export.
Note `CoInitializeEx` may return `S_FALSE`/`RPC_E_CHANGED_MODE` if the thread is already
initialized — only call `CoUninitialize` when the init actually succeeded (the
`shouldUninit` flag handles this).

## Commands you will need

| Purpose | Command | Expected |
|---|---|---|
| Build plugin | `.\build.ps1 -ProjectName ViewerImgRaw` | exit 0, 0 warnings/errors |
| Full build | `.\build.ps1` | exit 0 |
| Tests | `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` | all pass |

## Scope

**In scope**: `Plugins/ViewerImgRaw/ViewerImgRaw.Decode.cpp` (the production fix — this is the
**only** production file touched) and `Tests/ViewerPETests/ViewerPETests.cpp` (the regression
test + a test-only `IHost`/`IHostAlerts` stub).
**Out of scope**: `ViewerImgRaw.Export.cpp` (already correct); the decode logic itself; any
production code other than the COM-init block.

## Git workflow

- Branch: `advisor/015-viewerimgraw-wic-coinit`
- Message: `fix(ViewerImgRaw): initialize COM on decode workers so WIC formats decode`
- Do NOT push/PR unless instructed.

## Steps

### Step 1: Initialize COM inside `DecodeImageToBgraWic`

Put the COM init at the **single chokepoint** — the top of `DecodeImageToBgraWic`, before the
`CoCreateInstance` calls — so all three call sites are covered at once and any future caller
is too. Insert the export pattern:

```cpp
HRESULT DecodeImageToBgraWic(...) noexcept
{
    outWidth = 0; outHeight = 0; outBgra.clear();
    if (! data || sizeBytes == 0 || sizeBytes > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
        return E_INVALIDARG;

    const HRESULT coinitHr  = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    const bool shouldUninit = SUCCEEDED(coinitHr);
    auto coUninit           = wil::scope_exit([&] { if (shouldUninit) { CoUninitialize(); } });
    if (FAILED(coinitHr) && coinitHr != RPC_E_CHANGED_MODE)
        return coinitHr;

    wil::com_ptr<IWICImagingFactory> factory;
    HRESULT hr = CoCreateInstance(...);    // existing
    // ... rest unchanged ...
}
```

No include change is needed: `<wil/resource.h>` (for `wil::scope_exit`) is already available
transitively via `ViewerImgRaw.h` (which `#include`s it), and `wil::scope_exit` is **already
used** in this file (e.g. `:1076`, `:1276`, `:2015`, `:2368`, `:2376`). The fix compiles as-is.
Insert the COM-init block between the `E_INVALIDARG` guard (`:670-673`) and the first
`CoCreateInstance` (`:675-676`).

**Verify**: `.\build.ps1 -ProjectName ViewerImgRaw` → exit 0, no new warnings.

### Step 2: Confirm no double-init regression

Because `DecodeImageToBgraWic` is the only WIC entry on the decode side, do NOT also add
`CoInitializeEx` to the three call-site lambdas — one init at the chokepoint is enough and
nesting init/uninit per call is wasteful. Verify the three lambdas (`:2100`, `:2619`,
`:2751`) were left unchanged except that they now succeed.

**Verify**: `.\build.ps1` → exit 0.

## Test plan

> **Review correction (2026-06-21):** the obvious assertion — "a visible child window is
> present" (`CountVisibleChildWindows >= 1`) — **cannot detect this bug**. ViewerImgRaw renders
> the decoded image *and* the "no image" status via Direct2D (no child HWND for content); the
> only visible children are the DxUi combo-host chrome, which appear whether or not decode
> succeeds (see the existing JPEG case reusing that assertion at `ViewerPETests.cpp:2106`). A
> PNG test built on it would be GREEN with the bug present — worthless as a regression test.
> The decode result must be made **observable** through a real signal.
>
> Also note **why the bug went undetected**: the existing model case
> `TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo` (`:1996`) opens **JPEG**
> (turbojpeg — never touches COM), so it was always green and never exercised the WIC path.
> A *new* PNG case is required, not a tweak to the JPEG one.

Add an isolated case to `Tests/ViewerPETests/ViewerPETests.cpp` that observes decode success
via the **host alert** path (zero production change — production diff stays exactly
`ViewerImgRaw.Decode.cpp`):

1. **Fixture** — add `WriteTinyPngFile()` next to `WriteTinyJpegFile()` (`:1217`) using a
   committed, WIC-verified 120-byte 1×1 PNG byte array (generated via GDI+ and round-tripped
   through the WIC-backed `PngBitmapDecoder` → confirmed `1×1 Bgra32`), written via
   `WriteBinaryFile(path, std::as_bytes(std::span(kTinyPng)))`. No binary asset committed.
2. **Host stub** — add a small stack `IHost`+`IHostAlerts` stub (mirror `BuiltinFileSystemStub`'s
   refcount idiom: `std::atomic<ULONG> _refCount{1}`, **no `delete this`**). It records, with
   atomics, every `ShowAlert` whose `severity == HOST_ALERT_WARNING` and every `ClearAlert`.
   `_hostAlerts` is wired purely via `host->QueryInterface(IHostAlerts)` (`ViewerImgRaw.cpp:587`),
   so passing this stub as the factory's `host` arg makes the decode outcome observable.
3. **Drive** — model on `:1996`: LoadLibraryEx `ViewerImgRaw.dll`, `RedSalamanderCreate`,
   `createFn(__uuidof(IViewer), &factoryOptions, &hostStub, kViewerImgRawPluginId, …)`. Open a
   **single** PNG (`focusedPath` = the PNG, `otherFiles` = just that one → deterministic, no
   neighbor prefetch) via a `BuiltinFileSystemStub`. Wait for the window, then `PumpUntil`
   until a terminal decode signal is seen (`warningAlerts > 0 || clearAlerts > 0`, 8 s timeout).
4. **Assert** (race-free — all three alert sites fire only at decode *completion*, never at
   open-start: `ClearAlert` at Decode.cpp `:1485`/`:2912`, `ShowAlert(WARNING)` at `:2949`):
   - decode completed (a terminal signal was observed within the timeout), **and**
   - `warningAlerts == 0` (no decode-failure alert) → PNG decoded successfully.
   On the buggy code `DecodeImageToBgraWic` returns `CO_E_NOTINITIALIZED` → the async-failure
   path fires `ShowAlert(HOST_ALERT_WARNING)` → `warningAlerts == 1` → the case **fails**.
5. **Register** exactly like the model case: add to the `isolatedTests` vector
   (`RunFullSuiteInFreshProcesses`, ~`:4620`) and add a `shouldRun(...)` dispatcher block in
   `wmain` (~`:4681`).

**Prove the regression value (mandatory — this is the plan's #1 risk):** build with the test
but *before* the Step-1 fix and confirm the new case **FAILS**; then apply the fix, rebuild
`ViewerImgRaw`, and confirm it **PASSES**. A test that passes on unfixed code gives false
confidence.

- Verification: `.\build.ps1` exit 0, then `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild`
  → all pass, including the new PNG case.
- GIF/BMP/TIFF/HEIC share the same WIC path and are fixed by the same change but remain
  covered-by-inspection only (acceptable for a P1/S fix).

## Done criteria

ALL must hold:

- [x] `.\build.ps1 -ProjectName ViewerImgRaw` exits 0, 0 warnings/0 errors. ✓ (2026-06-21)
- [x] `.\build.ps1` exits 0. ✓ (full solution, 0 warnings)
- [x] `DecodeImageToBgraWic` calls `CoInitializeEx`/`CoUninitialize` (paired via
      `wil::scope_exit`, guarded by a success flag) around its `CoCreateInstance`. ✓
- [x] A test opens a PNG in ViewerImgRaw and observes it decodes (asserts no
      `HOST_ALERT_WARNING` fires), and that test **fails on pre-fix code / passes on fixed code**.
      ✓ Empirically proven: `TestViewerImgRawDecodesPngThroughWicWithoutErrorAlert` FAILS on the
      pre-fix DLL (`[FAIL] …without a decode-failure alert`, exit 1) and PASSES on the fixed DLL
      (all 11 checks, exit 0).
- [~] `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` passes. **Plan-015 scope is fully green**
      — the new PNG case passes inside the full suite, all ViewerPETests pass, all plugin DLL
      debug-selftests pass, and PluginContractTests passes (incl. `Create(builtin/file-system-7z)`).
      The full run reports 5 **pre-existing, unrelated** failures (this change only touches a plugin
      DLL + a test, so `RedSalamander.exe`/`FileSystem7z.dll`/monitor/Tools binaries are
      baseline-equivalent): (1,2) two `[CompareDirectories]` 7z cases fail with hr=0x80070490 —
      root cause is the stale `--plugin-path` selftest-harness option (logged "option --plugin-path
      no longer exists"; the 7z factory itself is proven healthy by PluginContractTests); (3)
      `search_service_sqlite_startup_warms_overridden_roots` is flaky timing (PASSES in isolation,
      exit 0); (4) `RedSalamanderMonitorEtwLatency` and (5) `ToolsPesterTests` are environment/
      tooling. None involve ViewerImgRaw or the WIC decode path.
- [x] The only **production** file modified is `ViewerImgRaw.Decode.cpp`; the only other change
      is the regression test in `Tests/ViewerPETests/ViewerPETests.cpp`. ✓
- [x] `plans/README.md` updated. ✓

## STOP conditions

- Excerpts don't match live code (drift).
- `DecodeImageToBgraWic` turns out to be reachable from the UI thread / an STA thread where
  `COINIT_MULTITHREADED` would return `RPC_E_CHANGED_MODE` and you cannot proceed without it
  — in that case the guard already returns S-path on `RPC_E_CHANGED_MODE` (treat it as
  "already initialized, proceed"); only stop if WIC then still fails.
- The new PNG test fails to decode even after the COM-init fix (indicates a second, separate
  defect — report it).

## Maintenance notes

- Any new background worker that uses COM (WIC/Shell/Direct2D device creation off the UI
  thread) must initialize COM per-thread; this is the project convention (see
  `Export.cpp:723` and `Common/.../IconCache`). Consider a tiny shared RAII helper if a
  third site appears.
- Reviewer: verify `CoUninitialize` only runs when init succeeded, and that the chokepoint
  approach (not per-call-site) was used.
