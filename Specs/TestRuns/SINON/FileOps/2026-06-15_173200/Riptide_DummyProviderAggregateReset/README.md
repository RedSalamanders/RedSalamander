# Riptide Dummy Provider Aggregate Reset Evidence

Date: 2026-06-15
Branch: `claude/zealous-poitras-ce383a`

## Root Cause

The post-R1-1 broad FileOps RED was an aggregate selftest harness state leak, not a product FileOps regression. Four broad failures passed as exact standalone cases. The archived broad trace showed the first real symptom at `Phase5_PreCalcCancelReleasesSlot`; later `Phase10_PermanentDelete`, `Phase11_ConnectionOverrideClamp`, and `Phase13_PostMortemDiagnostics` failures were downstream contamination.

`SelfTestState::dummyPaths` was intentionally retained across aggregate families, while cleanup intentionally skipped restoring FileSystemDummy configuration. Setup claimed the next family reseeded dummy deterministically, but only actually did that when `dummyPaths.empty()`. Later families therefore used stale dummy source paths and missing destination folders after intervening dummy provider config changes.

## Fix

`FolderWindow.FileOperations.SelfTest.cpp` now remembers the selected dummy seed config, reapplies it at each family setup, and recreates the required dummy destination folders without rediscovering the source path on every family transition.

## Evidence

- RED: `RED_post_R1-1_broad_dummy_state_leak/` copied from the post-R1-1 broad FileOps RED archive; 77 passed, 4 failed, 24 skipped.
- Build: `msbuild-20260615_171320_487.log`; 0 warnings, 0 errors.
- Focused GREEN: `GREEN_Phase5_PreCalcCancelReleasesSlot/`; 3 passed, 0 failed.
- Family GREEN: `GREEN_FileOpsFamily_Phase05_PreCalc/`; 9 passed, 0 failed.
- Broad GREEN: `GREEN_broad_FileOps_after_dummy_reseed/`; 85 passed, 0 failed, 20 expected skips, wall time 14m37s.
