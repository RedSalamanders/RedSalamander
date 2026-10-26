# Cinderstar package 20 closeout — pre-commit and publication verification

- Date: 2026-08-09
- Machine: `4cb089111a23`
- Build: x64 Debug, branch `operation-cinderstar-review`
- Build log: `.build/logs/msbuild-20260809_180720_511.log` (0 warnings, 0 errors)

## Results

| Scenario | Result | Duration | Key evidence |
|---|---:|---:|---|
| `Causeway_BridgeProviderOutputContracts` | 3/3 pass | 516 ms case | all four provider violations return `ERROR_PARTIAL_COPY`; write-underconsume is rejected before publication and leaves no final or staging file |
| `Cinderstar_LegacyWriter*` | 7/7 pass | 781 ms across five cases | eventual miss uses 2 probes; permanent miss 3; wrong size/replacement/cancel 1 each; post-publication objects remain fail-safe |
| `Floodgate_CrossFsMoveCommitSizeProofAvoidsImmediateRestat` | 3/3 pass | 1,641 ms case | 16 accepted proofs, 0 fallbacks, 0 immediate probes, 16 mandatory cleanup probes |

The closeout regression run first exposed that a writer could persist one byte fewer than it reported, commit the object, and force the new identity-safe post-publication verifier to retain the corrupted final. The bridge now verifies writer-local `GetPosition` against the pump count before `Commit`; release-without-commit retains operation-owned abort semantics. The legacy retry and concurrent-replacement behavior remains unchanged.

`causeway-precommit/`, `legacy-reverify/`, and `proof-control/` each retain unified-runner results, focused FileOps traces, case history, dashboard, and the complete performance stream. The archive contains 31 files and remains below the per-file and per-run size limits.
