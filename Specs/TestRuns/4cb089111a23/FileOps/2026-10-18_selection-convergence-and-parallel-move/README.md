# Copy/Move selection convergence and parallel managed Move

Same-machine x64 acceptance on 2026-10-18, base commit `a48e7d7880cb37ea6898d462ae188a6213a5f3f8` plus the dirty source snapshots bound by the retained build receipts. Existing unrelated worktree changes were preserved. Debug runs use `Z:\src\RedSalamander`; Release uses the attached isolated checkout `C:\Users\eric\.codex\worktrees\selection-shape-release\RedSalamander` because the user's independently running ordinary Release executable is occupied. No user process was terminated and no foreground input was used.

## Scope and fixtures

The governed runner uses marked `Z:\RedSalamander.Perf`, with raw artifacts preserved under `evidence/selection-convergence/<run-id>/artifacts` before the next runner's stale-run cleanup. Real SMB uses the enabled `path.smb.primary` resource `\\rubygloom\RedSalamander-SelfTest\selftest`, unique harness-owned child directories only. Local sources and independently generated expected contents remain in the sandbox.

Each matrix row copies or moves eight distinct 4 MiB files and one zero-byte file, once as a folder and once as the equivalent selected paths. The 8 MiB/s task limit makes overlap observable. Local Copy controls Manual four/one and Auto four with inactive saved Manual one. SMB covers native Copy, host bridge Copy and managed Move at four. Exact destination bytes and source outcomes are checked after timing. Bridge overlap counts file transactions holding I/O permits through publication/cleanup, excluding Opening rows waiting for a permit. Native overlap observes actual partial payload/finalization/publication work. Live selected-file popup snapshots must show determinate progress, exact totals and closed discovery.

Additional candidate controls cover 4,097 selected files (4,096 empty and an 8 MiB payload), frozen-total duplicate/changed-size protocol reports, cancel, a new unselected source child, known cleanup Retry/Skip, unknown cleanup outcome without replay, and a bounded thirteen-file/six-directory nested SMB Move with empty directories and independent content/timestamp checks. Protocol controls are synthetic reports on a real counted task, not a claim of a physical source mutation.

## Baselines

- `baseline-local`: run `20261018T133450Z-102188-1e5639da9cd84f9c85f80fabd8031c94`, full Debug build receipt `dfd02b4474c9c3902b0773123a6168d87c1fdd5e00ad081688d30fe1ab26128f`. Auto folder Copy resolved four but ran one; equivalent selected files ran four. Selected-file discovery stayed open 5.328–5.437 seconds while 29–32 MiB had already transferred. Manual four and one remained their respective budgets. All output contents and sources matched.
- `baseline-smb`: run `20261018T125900Z-29468-440d2b8db20b4c56948f42fabb0cb2dd`, targeted Debug build receipt `5cb7ddc22e2be26125e7566ea8814e1b14a5859009b1600635f1dad8b95c5953`. Folder managed Move reached one, selected-file Move four. Native/bridge Copy already reached four with Manual four. Selected-file discovery closed after all 32 MiB had transferred (9.969/13.484/14.062 seconds respectively). Folder/selected-file Move elapsed 34.250/19.342 seconds. Every content/source check passed.

Both baseline cases intentionally failed behavioral acceptance and exposed a fixture-only normalized-configuration string comparison, corrected in the candidate. Earlier observer experiments are retained only in the raw sandbox evidence; partial-byte-only observation incorrectly reported zero for bridge single-chunk files and is not used as a concurrency baseline. Same-machine timings remain subject to SMB latency; concurrency and early correct counting are the primary predicates.

## Evidence interpretation

`results.json`, `trace.txt` and runner aggregates are unmodified copies. `selection-metrics.jsonl` contains explicitly filtered, unmodified metric rows relevant to these predicates; the full metrics remain in the raw sandbox archives. Build receipts include source snapshot, profile, toolchain/dependency and output identities. A zero discovery duration in task metrics uses millisecond tick resolution; the separate fresh-metadata metric has microsecond resolution.

The sandbox audit reports 41 preexisting, unowned historical run directories. They were preserved; this does not establish a clean sandbox audit. Focused results do not qualify Full, foreground input, ARM64 or ASan. Candidate outcomes and final archive validation are recorded below after execution.

## Candidate qualification before the final local-drive/publication guard

| Profile/scenario | Governed run | Result |
|---|---|---|
| Debug local, exact current counting/cleanup fixture | `20261018T151723Z-72804-1e10d66f73d24e09b30064eb9bb7021c` | All six shape/concurrency rows, live totals and 4,097-file convergence controls passed. |
| Debug real SMB | `20261018T144040Z-65212-10f906fe629842d683ccd684d6e01009` | All six rows at four actual admitted files, all five cleanup controls, and bounded nested Move passed; no skips. |
| Release local | `20261018T151915Z-72804-b47b30de57134680afaf42094f408c59` | Six shape/concurrency rows, live totals and wide fresh counting passed; no skips. Earlier Release run `20261018T143253Z-103972-0296a81358874bfcb872ed49006ec47c` also closed 4,097-file discovery in 15 ms before payload bytes. |
| Release real SMB | `20261018T150546Z-64152-02235241de524e888d7cc640fcd78821` | Same complete focused SMB scope passed; no skips. |
| Debug remediation | `20261018T151030Z-25244-86f8534ebec44b18a1f735a6176a44c0` | `FileOps_ReviewRemediationContracts` passed. |
| Debug small-file cost | `20261018T151159Z-25244-877307b7a8744d299076fb9c5e376fe4` | `FileOps_SmallFileCopyCost` passed; treatment metrics retained. |
| Debug multiple folders | `20261018T151254Z-25244-13135c0a43024a25b4ddf344ff1b9776` | `DiscoveryScope_MultipleFoldersLiveProgress` passed. |

The Debug matrix above binds receipt `929c559c88d3ffcb31979339e063b0302980377c2fa95a452ac6bea2f7df0d77`; Release binds `0240b08168b90d592d58bff221a0cf96644e0da55e769d2a5fcb7e5d5d9eea94`. Both full-solution builds completed with zero warnings/errors. Exact per-run receipts are included where noted; the current source's subsequent scheduling/publication guard requires its own builds and runs below.

The NAS refuses exact directory metadata rebind (`ERROR_NOT_SUPPORTED`, 50). The nested scenario verifies `metadata.basic.unsupported` when folder times cannot be restored. It does not claim preserved timestamps on this resource, and it never changes an unproved replacement directory merely to restore metadata.

## Failed controls and wider qualification limits

- The first candidate SMB run caught missing Move byte-certainty publication to the popup despite correct engine totals; its raw run is preserved in `candidate-debug/smb-ui-failure` and the sandbox evidence.
- Run `20261018T140042Z-102436-50a64d8401204965beb6d961cbae33d2` fired cancellation with four I/O permits but only one confirmed cleanup waiter. Another file was still publishing; one staging file remained. That failed predicate remains under `candidate-debug/smb-checkpoint-failure`. The corrected fixture holds four actual post-publication cleanup waiters and then requires exact intact sources, complete destination bytes and no staging debris.
- That same run attempted the separate legacy 350-file/two-selected-root SMB workload. It copied 165 payloads before the existing 300-second deadline, then canceled and drained. This larger cost workload remains unqualified; the bounded thirteen-file/six-directory acceptance case does not replace or claim that gate. NAS latency varied substantially between runs, including native Copy controls whose concurrency was unchanged. Elapsed treatment timings therefore do not establish a causal throughput multiplier.
- Run `20261018T142428Z-80296-1fdb19762d0b4c4b83a8cd590d2864a6` passed the six rows and all five safety controls but failed an unconditional timestamp assertion. Its raw directory was removed by the next runner before preservation. `candidate-debug/smb-metadata-contract-failure/transcript-evidence.txt` explicitly retains only a transcript-derived summary and the build receipt; no missing raw artifact is represented as preserved or passing.
- One Release build failed C1090 (PDB RPC unavailable). `candidate-release/failed-pdb-build.json` retains its error and log hash. A repeat with unchanged sources built successfully; this is a toolchain recovery, not a source-level fix or runtime pass.

## Broader PR accounting

Debug PR run `20261018T144541Z-89720-a6dc70261323448495111a7a480c5dde` executed 13 entries. All twelve native/standalone entries passed. Pester passed 951 tests and failed one: `keeps historical workflows frozen behind their explicit archival triggers`. The Copy/Move exact no-activation admission, repair controls, phase inventory and I42 exact Done-path/sibling-rejection cases passed; their named results are retained.

The failure is already present in clean `HEAD` for `.github/workflows/terminal-engine-gate0.yml`: expected SHA-256 `e6f15b6fef9c0d367e12dc7f43cb769106dc982934658f501120728599f32d1c`, actual `cc43e0fdd4f239e6623ab9cd95837c4c0ed4ef864cf29400822a62eac94a5e64`. Compared with sealed commit `3e84e87b7`, the workflow has dependency-bot upload/download action version bumps. This task changes neither the frozen workflow nor its integrity guard. `pr-debug` preserves the failed aggregate, exact plan/decision, source identity and relevant named results. PR is **failed**, not green; the repository remains **NOT_EVALUATED** for Full. The PrePush explanation did not establish remote CI equivalence and retained five interactive obligations. No CI success, Full, ARM64, ASan or foreground-input pass is inferred.

## Local-drive and early-publication guard

The final implementation restricts eager selected-file metadata to Local fixed/removable/RAM disks. It caches only drive classification, never selected metadata. Network sources keep their existing producer I/O scheduling. Fresh byte certainty is published before closing/notifying the last leaf, rather than depending on a later payload callback.

Debug full-solution receipt `dec2240aea503dfcb6b732bb88107072c6c77ae5877da8b4250cbb8ae5e99daf` built with zero warnings/errors. Run `20261018T153747Z-79532-aca24049aa4a4ad0b134e8d8a25f5a27` passed all local rows and controls; 4,097-file discovery closed in 47 ms with zero payload bytes and a valid live denominator. Run `20261018T153952Z-79532-a68fb79a7aca40e299ccff81328953a7` passed the full focused real-SMB scope: all six rows reached four active files, exact contents/source outcomes passed, selected-leaf discovery closed before any payload bytes, five cleanup controls passed, and nested Move enumerated each of its six directories exactly once. `guard-debug` retains both raw result/trace copies, relevant unmodified metric rows and the exact receipt. Release guard compilation receipt `e28a125e5ea3b790f108e3c5938e5dddbb5fbf80581650e96d138b9e6ae7f2e0` also completed with zero warnings/errors, before the final plan move.

## Final post-closeout source validation

The completed I42 plan is under `Specs/Plans/Done/`; the authoritative contracts are updated. Final builds/runs after that document move are recorded here with matching receipts. Earlier passes above remain historical evidence with their original source identities.

Final source checks: `git diff --check` passed; specification inventory reported zero blocking findings; all 197 inventoried tool files are classified with zero findings. The explicit evidence check passed for 87 files before final-run additions, and the whole-inventory archive check passed for 13,844 files. A final explicit check covers the added final-run files below. The whole-inventory check was read-only and did not alter earlier evidence.

| Final profile/case | Run | Qualification |
|---|---|---|
| Debug local | `20261018T155203Z-30364-53e68bf229124d06aaf94c465320c992` | Passed all six matrix rows, live totals and wide selection controls. Four-file rows reached four; Manual-one controls reached one. 4,097-file discovery closed in 78 ms before payload bytes. |
| Debug real SMB | `20261018T155415Z-30364-759ae1ac58a44b6cae5b8ca8b3e55a0a` | Passed all six rows, five cleanup controls and nested Move. No failed/skipped cases. |
| Release local | `20261018T160016Z-107956-b55d67575dc84c2db512ec01212efabd` | Passed all six matrix rows, live totals and wide selection controls. 4,097-file discovery closed in 15 ms before payload bytes. No failed/skipped cases. |
| Release real SMB | `20261018T160205Z-107956-8e80fb9c1d1443bca5e65f0393ed700f` | Passed all six rows at four actual admitted files, exact contents/source outcomes, five cleanup controls and nested Move. No failed/skipped cases. |

The two final Debug runs bind full-solution Debug receipt `4314c81ad38d2c0631779fa4eff9ccc1d82811292e500155c770ef4b8347fc38`, source snapshot `c76baa325f6bcd0a1271dd67ebd150143cd1aec8b9cda59054af1f0ddae3041a`; the build had zero warnings/errors. `final-debug/<case>/qualification.json` links each result to its receipt/source/raw archive. Final Debug runtime ran while the isolated Release profile compiled; its elapsed transfer times are not a causal throughput comparison.

Final Release binds full-solution receipt `27345aa8460edbe4db24d92b4ef663d4d28b6f7f8175bb1ba4b8b1d5a42a0471`, snapshot `07f633e4f279c81d677a4b949c195c6e2070b16384b48b59c184758ff195e5fd`; the build had zero warnings/errors. The snapshot digest includes physical checkout identity, so it differs between the two roots. `final-source-comparison.json` independently confirms identical HEAD/base and all seventy source-file records, including content hashes, modes and dirty status. `final-release/<case>/qualification.json` binds each Release result to its own receipt/source/raw archive.

All four final governed runs passed with zero skipped cases and used no desktop activation. Selected-leaf matrix totals closed before payload bytes and never grew after closure. Four-file scenarios reached four active transactions; explicit-one local controls reached one. The completed scope is Copy/Move selection-shape consistency and fresh local-disk selection convergence, with the broader qualification limits above preserved. The existing ordinary Release process was not replaced or stopped; the tested updated Release executable is in the attached `selection-shape-release` checkout.

Final explicit archive validation passed for all 113 files (approximately 2.68 MB, below the 5 MiB run budget); `git diff --check` also passed. Both final SMB metric sets were independently checked for six four-file/content/source/closed-total rows and five four-waiter cleanup controls. The earlier whole-inventory result and the unrelated failed PR result retain their original scope.
