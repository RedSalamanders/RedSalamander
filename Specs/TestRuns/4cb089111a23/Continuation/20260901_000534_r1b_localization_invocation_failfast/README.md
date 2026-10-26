# R1b LocalizationTests invocation fail-fast

## Purpose

This compact record preserves the failed direct `LocalizationTests.exe`
invocation before the corrected focused rerun. The dump remains in its original
local location and is intentionally not copied into the repository.

## Failed invocation

- Source commit: `b9df3a3a21d90d2c109cb5fc1dcf3b5d98f7e116`.
- Executable: `.build\x64\Release\LocalizationTests.exe` from test-enabled
  full-solution receipt
  `38841fe2388d7d8522fa0cb0c7d5a7b3d1c0f211623b4780c626f85367973024`.
- Environment set `REDSALAMANDER_TEST_ROOT=D:\RedSalamander.Perf` but omitted
  `REDSALAMANDER_TEST_RUN_ID`.
- Result: the resource/satellite cases passed; all settings-file writes failed;
  the final unchecked `optional::value()` in the `noexcept` settings test
  terminated the process with `0xC0000409` (`FAST_FAIL_FATAL_APP_EXIT`).
- This is an invalid test invocation, not R1b product evidence. Test-enabled
  `SettingsStore` deliberately returns an empty settings path when a configured
  unified test root lacks a safe run ID, so it cannot fall back to user data.

## Crash identity

- Process ID: `70544`.
- WER event name: `BEX64`; fault module `ucrtbase.dll`; exception
  `0xC0000409`; subcode `7`.
- WER report ID: `74f8d215-180a-4241-9da1-cfefa207938b`.
- WER archive:
  `C:\ProgramData\Microsoft\Windows\WER\ReportArchive\AppCrash_LocalizationTest_9a48a5989b66441f1752554d9b37c134849bc5_90f00c3a_57acef88-00ac-42ef-8fff-32a10b2e8d56`.
- Dump: `C:\Users\eric\AppData\Local\CrashDumps\LocalizationTests.exe.70544.dmp`.
- Dump size: `610,920` bytes.
- Dump SHA-256:
  `171E56549F73DAC3590F8F1A6C8FC88602147FC8BF31D1825CCBF1F246F9BE36`.
- No RedSalamander sidecar stack was produced.

## Continuation

Rerun the exact Release binary with both unified sandbox variables set, using a
fresh safe run ID. Count only that corrected invocation in the R1b closeout.
