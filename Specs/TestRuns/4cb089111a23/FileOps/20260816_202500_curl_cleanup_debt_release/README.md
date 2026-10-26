# Curl structured cleanup-debt validation

- Date: 2026-08-16
- Machine profile: `4cb089111a23`
- Source base at capture: `e9c76e59aa2b2993170ea1e6c88e7d15362954c7` plus the uncommitted Package D remediation under test
- Build: test-enabled x64 Release
- FileSystemCurl build: `.build/logs/msbuild-20260816_202715_439-pid3768-a1cca632.log` — 0 warnings, 0 errors; receipt `77383e4563a8e026c8fbdbaa1bc03744a4ed105320d4da5898d28b8bc9a03bb9`
- PluginContractTests build: `.build/logs/msbuild-20260816_202805_502-pid3768-76dc10a3.log` — 0 warnings, 0 errors; receipt `2248d99f39c924b815dd9781beab9b03ac470b5398c518ed65f354d73fc5e391`
- Focused entry point: `.build\x64\Release\PluginContractTests.exe --curl-selftests`
- Result: five separate processes, each 221 passed / 0 failed with a clean FileSystemCurl unload quiet point
- Raw bounded metrics: `perf_metrics.jsonl` (350 valid JSONL rows, 179,409 bytes, SHA-256 `4f8cb519d6de71dd727cbf4e8c0b3311000590434da10958ff1d193968008247`)

The deterministic loopback matrix protects the structured publication/recovery result across writer, single, recursive, and batch Copy/Move paths. It distinguishes committed publication, primary failure, cleanup failure, and source-deletion failure; retains rollback, staging, and uncertain tree artifacts without unsafe pathname deletion; aggregates cleanup debt once per public operation; keeps callbacks and normal notifications truthful; and proves that a second successful writer `Commit()` is idempotent. The failure controls include deterministic upload rejection after storage, exact-size verification failure, prepare/promotion failure, and failed safe cleanup of an operation-owned staging sibling.

## Bounded evidence

| Metric | Five Release samples | Median | Stable work/result |
|---|---|---:|---|
| `FileOps.Curl.CleanupDebtGuard` duration | 60,000; 60,000; 59,000; 59,000; 62,000 us | 60,000 us | Six committed overwrite shapes, exactly 320 deterministic loopback commands and nine retained rollback items per run, `S_OK` |
| `FileOps.Curl.CleanupDebt` observation rows | 19; 19; 19; 19; 19 | 19 | Exactly 22 retained-item counts aggregated per process; one fixed content-free detail and no endpoint/path/user/credential/payload text |

The 19 cleanup-debt observations per process cover all cleanup-debt-producing Curl scenarios in the focused suite, including preservation controls from the accepted conditional-publication matrix. Their deterministic distribution per process is: eleven single retained-rollback observations, three two-item retained-rollback observations, two retained-staging observations, and three preserved-destination-tree observations. The aggregate retained-item count is therefore 22 per process. The metric detail is always `aggregate retained remote recovery items`; a forbidden-content scan found no endpoint, authority, path, randomized sibling, password, or credential marker.

This archive establishes a same-machine bounded correctness and safety-cost baseline. No apples-to-apples historical metric exists for the cleanup-debt guards, so it makes no improvement or live-endpoint throughput claim. The earlier source-commitment and conditional-publication archives protect different behavior and are retained as independent baselines.
