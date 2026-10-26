# E6 post-retirement artifact projection candidate

- Machine profile: `4cb089111a23`.
- Source: HEAD `9ae7c0cff9f657c42157f1913d9d54edff2bfefd` plus the
  uncommitted E6 evidence-honesty correction before a source-guard-only
  synchronization.
- Performance build: test-enabled x64 Release full solution; artifact receipt
  `1518a7903f2231f88029766c5afc8633d579d057a26587b29c55f0019b980371`;
  diagnostics `0 warning(s), 0 error(s)`.
- Command: `Run-AllTests.ps1 -Suite Commands -Configuration Release
  -SkipBuild -CaseFilter cmd_fileops_artifact_name_shape_projection_scaling`.
- Scenario: five independent processes project 10,000 ordinary names and 100
  recognized Possible artifact names through the production post-retirement
  name-shape classifier used by FolderView and Find Files.
- FolderView durations: `1281`, `1548`, `1297`, `1274`, and `1263` us;
  nearest-rank p95 `1548` us.
- Find Files durations: `1916`, `2190`, `1848`, `1986`, and `1819` us;
  nearest-rank p95 `2190` us.
- Regression gate: PASS. Every sample reports zero ordinary-row artifact
  probes (`value0=0`), exactly 100 Possible-name probes (`value1=100`), and
  `hr=0` for both consumers.

All five candidate processes pass (`passed=1, failed=0, skipped=0`). Each
archived `perf_metrics.jsonl` retains the two claim-bearing projection rows in
original order. The original runner JSONLs remain under the initialized test
root and are bound here as `(rows, bytes, SHA-256)`:

- sample01: `(318, 120142, 2A6EDE395AC0C33277BEC853460E2D92B95A0DE7EA705E5E341DB44E19E1795C)`
- sample02: `(318, 119908, 3B1A4F2FEBE13B8B449496E8E2F775EE0D5AC964CA55C40ED0EC7C1470F0A045)`
- sample03: `(318, 119917, CEB6CE104658AF12C84A7F4D60D802907CD7FFB054259BFC6EF4603ACFF2449C)`
- sample04: `(318, 120282, 8C96833E8447BECFEC9E9AF8CD59C3415F81247705528108A4BCBF59311D4F1A)`
- sample05: `(318, 120244, 428CA99F1C7F77B7FFEB624BCC6153B7971F4948530B223DA007DE33AE22C436)`

The focused runner reported increasing disk-audit issue counts (`178` through
`182`) from sibling inventory already beneath the explicitly initialized
`D:\RedSalamander.Perf` root. These focused runs are not claimed as a whole-root
cleanliness proof and produced no evidence of a write outside the authorized
test root.

After the five samples, the stale admission source guard was synchronized with
E5's central scheduled-Rename owner; no production or scaling-case code changed.
The resulting exact-source test-enabled x64 Release full solution built with zero
warnings/errors under receipt
`7e0e8e4fc2b1fe413d99213042dd7e1575daee1a95c601bfe2e7d17122200132`,
and both `file_operations_phase0_typed_contract_source_guard` and an exact-source
rerun of `cmd_fileops_artifact_name_shape_projection_scaling` passed `1/1`.
