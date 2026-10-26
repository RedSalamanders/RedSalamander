# MSVC 14.52 Release evidence

The original sizes.json is the retained test-enabled x64 Release comparison, with RSBuildEnableTests=true on both sides. Its main EXE grows from 36,492,288 to 38,948,864 bytes (6.73%). PE version-resource markers independently confirm that flavor; it is not a shipping-size claim.

The separate production-only comparison uses RSBuildEnableTests=false, BuildNumber 1, default compiler debug records and one unchanged source path on both sides. RedSalamander.exe grows from 8,099,840 to 8,401,920 bytes (3.73%). Both images have no test-build marker or reserved test exports. Exact production image manifests, hashes, build receipts and the policy patch are retained under the user-approved E:/RedSalamander.Perf/production-size-evidence/.

Both comparisons select MSVC 14.52.37025. No compiler-version speedup is claimed. Runtime measurements use the retained test-enabled PerformanceTests2 icon-enumeration fixture and do not establish complete production-application performance or native ARM64 acceptance.

[Settings](Settings.md) and the compressed original snapshots preserve the selected fields. full-tool-metadata.csv.gz additionally records every captured compiler/resource/MIDL/MASM/shader/link/archive metadata field, with exact project/profile/file mappings. Intrinsic file paths/timestamps are retained in raw snapshots and excluded from setting comparisons. Final qualification, every failure and paired runtime results are reported at E:/DxUi.Perf/reports/Msvc1452ReleaseOptimization_2026-10-23/CrossRepository.md; unsafe failed runs remain retained in place.
