# Riptide R1-6 - Early-Admission Thread Safety

Date: 2026-06-15
Branch: claude/zealous-poitras-ce383a

## Root Cause

`Task::ThreadMain()` is `noexcept`, but the early-admission COPY path constructed a `std::jthread` directly. If thread creation threw `std::system_error`, the exception would escape the `noexcept` worker and terminate the process. The same path also ran for COPY even when pre-calculation was disabled, creating a thread that immediately returned.

## Fix

- Gate early admission on `_operation == FILESYSTEM_COPY && _enablePreCalc`.
- Move pre-calculation thread creation into `Task::TryStartPreCalculationThread(...) noexcept`.
- Catch only `std::system_error`, log a warning diagnostic, and fall back to serial pre-calculation before transfer.
- Harden multi-worker pre-calculation fan-out as well: worker `std::jthread` creation failures now continue with the workers that started, or run serially if none started.
- Add `ENABLE_TESTS` hooks to force thread-start failure and count attempts.

## Evidence

- `RED_hook_not_honored_old_engine/`: `Riptide_EarlyAdmissionThreadStartFailureFallsBack` failed before the fix because the forced thread-start failure hook was not honored (`attempts=0`).
- `FINAL_GREEN_thread_start_failure_fallback/`: forced pre-calc thread-start failure completed through serial fallback.
- `FINAL_GREEN_disabled_precalc_no_thread/`: COPY with pre-calc disabled completed with zero pre-calc thread-start attempts.
- `FINAL_GREEN_existing_early_admission_overlap/`: existing normal early-admission overlap behavior still passes when pre-calc is enabled.
- `FINAL_GREEN_fairstream_family/`: `FileOpsFamily_Fairstream` passed, 26 passed / 0 failed / 0 skipped.
- `FINAL_GREEN_broad_FileOps_after_R1-6/`: broad FileOps passed, 87 passed / 0 failed / 20 skipped, wall time 14m39.0s. Skips are expected 7z/remote credential-gated cases.

Build log: `msbuild-20260615_180149_790.log` (0 warnings / 0 errors).
