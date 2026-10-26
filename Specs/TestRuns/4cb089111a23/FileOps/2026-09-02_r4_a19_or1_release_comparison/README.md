# R4-A19-OR1 Local Delete observation comparison

- Machine profile: `4cb089111a23`
- Baseline source: `14915a84ce8a2c7590464aed8fc3f04f56c5cdb0`
- Baseline build: test-enabled x64 Release full solution; receipt `6ed185899f884fc5a295a611405bae845ab15bc1f67687e595dccafacc450440`; diagnostics `0 warning(s), 0 error(s)`
- Candidate implementation: `0c7ac7a4`
- Candidate source used for the final samples: `ada031d3`
- Candidate build: test-enabled x64 Release full solution; receipt `64a3705dc274bb1a03b157d1bc3a51a922d2865abe84f57ed4b61fcf7c22c8d0`; diagnostics `0 warning(s), 0 error(s)`
- Command: `Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter R4A19_DiscoveryMeasurementFacts -Configuration Release -SkipBuild -TestRoot C:\RedSalamander.Perf -TimeoutMultiplier 4`
- Raw archives: adjacent `2026-09-02_r4_a19_or1_{baseline,candidate}_release_01` through `_05` directories

## Five-process results

| Side / scenario | Duration samples (us) | Duration p95 | First-mutation samples (us) | First-mutation p95 | Open-discovery mutations |
|---|---:|---:|---:|---:|---:|
| Baseline Copy, 128 x 256 KiB | 156000, 156000, 156000, 156000, 156000 | 156000 | 4260, 4422, 4480, 4508, 4669 | 4669 | 128, 128, 128, 128, 128 |
| Candidate Copy, 128 x 256 KiB | 156000, 156000, 157000, 157000, 188000 | 188000 | 4363, 4422, 4848, 4950, 5257 | 5257 | 128, 128, 128, 128, 128 |
| Baseline Delete, 1,024 x 1 KiB | 78000, 78000, 78000, 125000, 156000 | 156000 | 0, 0, 0, 0, 0 | 0 | 0, 0, 0, 0, 0 |
| Candidate Delete, 1,024 x 1 KiB | 78000, 78000, 78000, 79000, 94000 | 94000 | 68687, 69122, 71338, 77138, 91457 | 91457 | 1, 1, 1, 1, 1 |

The baseline processes are expected RED: each passed setup/cleanup, completed exact Copy and Delete mutations, and then failed the immutable Delete observation assertion because the host reported zero open-discovery mutations. Each candidate process passed `3/3` with no skips. Candidate Delete reported exactly 1 MiB completed while discovery remained open and zero provider progress callbacks in every process.

## Gates

- Copy duration: `188000 <= max(1.35 * 156000, 156000 + 250000) = 406000 us` — PASS.
- Copy first mutation: `5257 <= max(1.50 * 4669, 4669 + 100000) = 104669 us` and `< 1 second` — PASS.
- Delete duration: `94000 <= max(1.35 * 156000, 156000 + 250000) = 406000 us` — PASS.
- Delete first mutation: baseline had no observable fact; candidate `91457 <= 100000 us` and `< 1 second` — PASS.
- Candidate maximum ready queue was 7 for Copy and 128 for Delete, below 256 in all five processes.
- Candidate starvation count and traversal-limit hits were zero. Maximum discovery callback work was 125 us, maximum discovery lock wait was 29 us, and bridge stage-retained count was zero.

Result: **PASS**. The correction observes an already-proved exact Local permanent-Delete result; it does not publish a synthetic provider callback or change scheduling, provider mutation, authority, or UI behavior.
