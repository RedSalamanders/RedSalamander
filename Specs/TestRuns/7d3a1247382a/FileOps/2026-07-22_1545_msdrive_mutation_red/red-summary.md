# EG-MSD-01 mutation-time identity validation RED

- Date: 2026-07-22
- Machine profile: `7d3a1247382a`
- Build: test-enabled x64 Release
- Provider build: `.build/logs/msbuild-20260722_154325_642.log` — 0 warnings, 0 errors
- Harness build: `.build/logs/msbuild-20260722_154415_988.log` — 0 warnings, 0 errors
- Command: `.build\x64\Release\PluginContractTests.exe`
- Result: exit 1; Microsoft Drive exported selftests reported 167 passed / 2 failed

The deterministic fake-Graph hook returned the original destination metadata and then moved that exact item to
`/elsewhere/important.txt`, changing its content and eTag before overwrite-backup preparation. The setup/result
assertion passed. The two expected failure-first assertions showed that the current production path did not preserve
the externally moved item/source and issued stale PATCH/DELETE mutations instead of one mutation-time validation GET
with no mutation. Production overwrite logic was unchanged for this checkpoint.
