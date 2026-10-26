# EG-S3-01 direct-publication race RED

- Date: 2026-07-22
- Machine profile: `7d3a1247382a`
- Build: test-enabled x64 Release
- Provider build: `.build/logs/msbuild-20260722_165809_115.log` — 0 warnings, 0 errors
- Harness build: `.build/logs/msbuild-20260722_170054_580.log` — 0 warnings, 0 errors
- Command: `.build\x64\Release\PluginContractTests.exe`
- Result: exit 1; S3 directory-transfer probes reported 16 passed / 2 failed

The deterministic fake S3 graph reports the destination missing during the whole-plan and just-in-time refreshes,
then creates `dest/race.txt` inside the next direct publication call. Production transfer logic is unchanged for
this checkpoint. The two expected failing assertions prove that no-overwrite COPY and MOVE both replace the
concurrent bytes; MOVE also deletes its source after that unconditional publication.
