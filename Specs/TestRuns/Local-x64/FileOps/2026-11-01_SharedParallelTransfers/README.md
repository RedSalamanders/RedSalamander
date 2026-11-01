# Shared parallel transfers — 2026-11-01

Implemented in `Z:\src\RedSalamander-worktrees\fileops-waiting-state-20261031`, branch `ej/fileops-waiting-state-20261031`, based on `6b6f5646be297fbc232d137f1df6a00c5cb74461`. Changes remain in that isolated worktree.

`fileOperations.maxParallelTransfers` defaults to four and accepts 1–16 in Preferences > File Operations > Advanced > Parallel transfers. Distinct bound regular sources with missing sibling destination names may transfer together when destination parent IDs are unsupported. The budget is shared across jobs and held through publication, verification and cleanup. Conflicting copy/move jobs retain serialization. Waiting cards explain capacity, exclusivity and identity waits, with no false Counting or Start now bypass. Delete and metadata-only native renames keep their own concurrency budgets.

The accepted test-enabled x64 Debug solution build reported zero warnings and zero errors:

- Receipt: `073b24253987538b8a13561df4a8a00b4570c538bce3e39c6e9bcf8cb0c3bb61`.
- Source snapshot: `78f132798a23e2d4df514355879d7b26e5120be234cfd41d19fbb14afc76d78c`.
- RedSalamander.exe SHA-256: `4c91ba5e1b80b7bbbe45eeacb163506d7b9dad39448e00ed7397a355a8893146`.
- MSVC 14.52.37025, supported Embedded `/Z7` Debug recovery mode. The earlier ordinary `/ZI` C1001 in unchanged `Commands.Tests.Runner.cpp` is retained in the adjacent identity-waiting archive. Ordinary `/ZI`, production Release, ARM64 and ASan qualification are not established here. Dependency pins were preserved.

Fresh focused results, with originals beneath `Z:\RedSalamander.Perf\evidence\completed-runs`:

| Evidence | Run ID | Result |
| --- | --- | --- |
| Root safety family, metrics and captures | `20261101T005245Z-81876-5d96eb249d5142babc1ca5f85060a9c6` | Seven selected cases plus Setup/Cleanup passed; zero failures/skips |
| `preferences/` | `20261101T005129Z-81876-e0409ed066624c92ad555c2738b66c6f` | Four selected settings/UI cases passed |
| `native-budget/` | `20261101T004910Z-81876-89c2e4c695194046a45471d72738fd26` | Native control passed; configured budgets four and two |
| `alternate-volume-skip/` | `20261101T005021Z-81876-fe01c42dd54845c1839646209c50bd37` | Alternate-volume case skipped: no authorized second test root |
| `conflicts/` | `20261101T003349Z-84020-6778aa2df0de4639981d8091f48b9a19` | Overlapping transfer serialization passed |
| `publication-faults/` | `20261101T003541Z-84020-556021d697cc4126816b01050ad811de` | Publication-failure/source-retention matrix passed |
| `small-files/` | `20261101T003713Z-84020-b01c15409d8445cdb759dd106fedf425` | Native and bridge 2,000-file controls passed exact output validation |

The last three are supporting results on receipt `738f8e46…`. `supporting-source-delta.json` compares their canonical snapshot with the final receipt: only the throughput test fixture and its testing specification differ; all production input bytes and Git HEAD match. Their original receipts remain explicit. This comparison does not promote exact-artifact reuse or establish a Full/PrePush pass.

See [scenario-contract.md](scenario-contract.md) for the six actual-I/O workloads, identity-injection capabilities, source-proof checks, 1/2/4/16 budget controls, canceled fifth job, multi-file jobs, alias/name conflicts and UI assertions. The same-candidate four-file Copy controls measured peak one at limit one (75,348 us) and peak four at limit four (39,669 us). Both include deliberate transaction barriers. The 64/4,096-file index witnesses used 896/57,344 probes (14 per selected file), below the 100-per-file bound, in 1,622/147,095 us. These are Debug scheduling observations; physical NAS throughput, native offload/block-clone behavior and an old-source/new-source performance improvement are not established.

The gallery's 24 identity/capacity/exclusive English/French Copy/Move scenarios produced 288 measured redraws; the review gallery contributed another 288. The first-pass `Show-PerfRuns.ps1` analyzer with `-FailOnQuality` exited zero for the combined 576 frames: median 8.205 ms, descriptive P95 8.471 ms, maximum 33.569 ms. P95 sample quality is sufficient; P99 is below its 1,000-sample minimum. This is no hard latency or throughput-improvement claim. Six inspected capacity/exclusive images are promoted. All captures use invented Studio/Moonrise paths and were produced by the harness without desktop takeover. Unpromoted identity frames retain the earlier intermittent missing-path painting observation; no path-painter fix is claimed.

The root metric stream is a digest-bound, original-order selection from the broad raw stream; `metric-selection.json` records original byte/row counts and line numbers. The small-file stream has its own selection metadata. Broad raw metrics remain with their exact original runs. `retained-failure-history.json` and the retained failure folders preserve earlier fixture, selection and implementation failures, including the corrected native-route fixture: individual new-file selections intentionally use the shared-permit host bridge, so its provider-native control now selects a directory root and verifies the nested corpus.

Focused native invocations report 32 pre-existing shared-sandbox audit issues. The broader PR run reports 34, including two newly observed BuildConsole fixture directories. Unrelated or unattributed paths were preserved; selected passes do not establish a clean whole-sandbox audit. The marked sandbox was already initialized. No alternate volume/root was created. The user explicitly authorized the needed focus-taking tests for the night; Preferences used the canonical runner's desktop lease, warning and restoration.

The broader Fresh local PR run `20261101T005544Z-81876-1556b6111196498d9231df9f38ac3b09` is retained in `local-pr/`: 1,083 passed, nine failed, two skipped. All eleven native/script entries passed; the tooling entry remains incomplete. Standalone schema tests passed, including limit round-trips/default pruning/type boundaries; UI, localization, plugin, viewer, crash, Curl and performance suites passed. This is not a green PR/Full result.

Four original tooling failures came from the isolated worktree lacking its pinned `.ci-engine` sources. After the run, an independent checkout of exact commit `c3c778cabd81dbf85b95f1dc4cec64b3b0eac12c` was restored without changing the other checkout. Source-only rechecks passed BuildReproducibility 23/23, ReleaseWorkflowPolicy 44/45 and WingetValidation 19/21. Remaining failures concern unchanged archived-plan, Dependabot, CI/nightly, Ghostty and Winget source expectations. Their names, stacks and original console output are retained, including failures formerly masked by missing-source Describe errors. Rechecks do not replace the original Fresh PR verdict. No workflow, publication or archived-plan source was changed for this feature.

The read-only closeout obligation plan records wider Full obligations and returns repository NOT_EVALUATED. Explicit archive and whole-inventory validation apply to this curated payload; neither establishes Full runtime qualification.
