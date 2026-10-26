# Affected x64 Release validation, 2026-10-18

Requested scope: repeat the affected Debug selections in Release, including real disk-to-SMB Copy
and Move; repair failures; update the authoritative contracts and retain the evidence.

## Tested candidate and scope

The checkout is based on master `6d2fcb04f85fb425fd9d83f2125af963341025e0`. The comparison base for
the delivered feature selection is `265147ee62c296b1a47ed44efb4d8edb58ad60cf`.
`candidate-scope.json` lists the exact eleven runtime selections, reasons and repeat counts.
The same three ViewerSpace cases and nine tooling files as the Debug packet are included.

The generic affected and PrePush explanations are retained in `explain-plan.json` and
`repair-prepush-plan.json`. Their conservative wider recommendations are advisory, not passing
results. The user explicitly requested the same affected tests; this packet does not claim Fresh
Full, ARM64, ASan, unrelated plugin/device coverage, the legacy 350-file/two-root SMB cost gate,
or clean-sandbox acceptance. The runner's 41 preexisting unowned historical run directories were
preserved. Hosted CI was blocked before this local validation by account billing/spending limits.

## Release failure and repair

The first test-enabled Release solution build had zero warnings/errors and receipt
`3e6b4068d6cfa867ec22b996e583bf298e41df53a21c747abb93ff90ce1551a5`, source snapshot
`d207a07fae17a201c04b87fccb587f382939ed9ed211f4dbe2dddd4a48eb5f0f`.
Seven selections passed before `FileOps_SmbManagedMoveLeavesNoStage` failed:
`20261018T201442Z-26568-5764ffaca95b46439a46a76a760478ae`.

Canceling the third Finalizing occurrence on a fresh NAS destination left the owned
`payload/nested/deeper.rs_tmp_...` directory. The original trace and metrics record an uncertain
publication and one retained stage. The first four cancellation phases passed. The full failed
results, trace, source/receipt identity and metrics remain unchanged; `baseline/` preserves the
original selection summary and source bindings.

Local conditional publication checked operation control before SMB metadata preparation, but did
not check again immediately before rename. The repair adds that last reversible checkpoint and
uses the existing exact rollback/owned-stage cleanup path. It neither suppresses the primary cancel
watch nor guesses at the result of an attempted rename. New diagnostics record a stopped
pre-rename boundary and the underlying Win32 error when rename reconciliation is actually unknown.

The existing SMB cancellation case now also exercises direct file and directory publication on
the registered target. A synchronous control cancels at its second check: publication must return
known non-commit, preserve the exact stage and absent final name, then remove the exact stage under
the bounded cleanup control. The original no-stage and content assertions remain unchanged.

The final test-enabled x64 Release solution rebuild also has zero warnings/errors:
receipt `02497e7e05d551155e4162728c24127e76c9b1246ca5988afe43fa56949bf928`,
source snapshot `8ebb9a33a6cd0384e9077f80f1edbbd872d7c1ae6ce6bea0473a2aed0352ff1b`.
`repair-source-delta.json` identifies one production C++ file, its native test fixture and two
authoritative spec changes relative to the failed binary. Whole-file formatting also normalizes
existing whitespace in the newly touched provider source.

All eleven selections were rerun against the repaired receipt. The previously failing case ran
three repetitions: six deterministic pre-rename controls and fifteen phase-cancellation trials.
All passed; retained-stage counters sum to zero and no rename-reconciliation unknown was emitted.
The new guard fired seven times, including one phase trial beyond the six direct probes.
The original phase failure is timing-dependent; these observations do not establish that every
already attempted network mutation can always be reconciled. Genuine unknown outcomes remain
conservative and are never replayed or cleaned by pathname.

## Completed selected validation

| Release selection | Passed | Failed / skipped |
|---|---:|---:|
| Managed SMB Move cancellation, three repetitions | 9 | 0 / 0 |
| Discovery scope, including painted denominator and capture checks | 27 | 0 / 0 |
| Eleven selected Commands cases | 11 | 0 / 0 |
| Decision mirror | 9 | 0 / 0 |
| Discovery cancellation/queue transitions | 11 | 0 / 0 |
| Bridge and connections | 9 | 0 / 0 |
| Native SMB concurrency budget | 3 | 0 / 0 |
| Same-provider bridge destination storage | 3 | 0 / 0 |
| Local selection shapes | 3 | 0 / 0 |
| Real SMB selection shapes and Move safety | 3 | 0 / 0 |
| Completed-card remediation | 3 | 0 / 0 |
| **Runtime total, including setup/cleanup** | **91** | **0 / 0** |

All three affected native ViewerSpace cases return zero with no activation, stable source and
receipt-verified artifacts. All 421 checks in the nine selected tooling files pass with zero
failures/skips; `RequiresBuildToolchain` remains explicitly excluded as in the Debug selection.
The tooling post-test barrier confirms both source identity and every attested artifact digest.

Actual resources: marked `Z:/RedSalamander.Perf`, the already marked alternate-drive
`D:/RedSalamander.Perf` enabled through the governed switch, and the configured
`\\rubygloom\RedSalamander-SelfTest\selftest` NAS test target. No share or mount was created.
The Commands focus cases used the harness warning and desktop lease. Other application captures
use the native no-activation harness; no manual desktop screenshots were used.
The NAS phase-cancel matrix uses fresh destinations; its identity-dependent overwrite submatrix
is intentionally available only on the loopback target, not claimed from this NAS run.

## Independently checked observations

`independent-metric-checks.json` reviews preserved raw JSONL after the native assertions:

- Local folder/leaf Copy pairs reach 4/4 files in Manual four and Auto with inactive Manual one;
  Manual one remains 1/1. All six rows verify exact bytes, source outcomes and frozen totals.
- Real disk-to-SMB native Copy, host-bridge Copy and managed Move each pass folder and selected-file
  rows: all six reach four admitted file transactions. Every row verifies nine files, 32 MiB,
  exact destination contents/source outcomes, no payload while discovery is open and no later growth.
  Selected-file rows also show determinate, closed live totals during transfer.
- Fresh metadata for 4,097 selected local files converges in the recorded 16 ms, before payload,
  below the unchanged two-second bound. The clock has tick resolution. Duplicate reports preserve
  the denominator; the changed-size protocol control withdraws certainty; cancellation preserves
  sources and leaves a clean destination. This is fresh enumeration, not pane-total reuse.
- All five parallel Move cleanup controls pass, including an injected unknown outcome without replay.
  The nested Move passes thirteen files, six directories, four-file overlap and post-order source
  removal. Unsupported destination timestamp metadata is explicitly noted rather than claimed preserved.
- The pipeline pair uses one 128 MiB file, 30 ms chunk latency and the existing buffers:
  110 ms setup outside both clocks, 1.375 s disabled, 0.891 s enabled. Exact payload and the unchanged
  candidate-at-most-80-percent threshold pass. This is one synthetic same-binary pair, not a live
  SMB throughput gain or a comparison with the unqualified historical seed-contaminated corpus.
- All four completed-card controls pass. The height-retaining counterfactual paints 427 px on its
  first frame; default, unknown-label and reduced-motion controls already paint 66 px (44 DIP at
  144 DPI), with zero first-frame timers and quiet final state.

SMB durations reflect a live endpoint and vary from the earlier Debug runs; no cross-profile or
unmatched before/after throughput claim is made. Final static checks pass: all 36 changed native
files satisfy pinned clang-format 22.1.3, whitespace checks pass, tool inventory is complete and
specification inventory reports zero blocking findings.

## Evidence and reproducibility

Use the standard test-enabled solution build and governed runner:

```powershell
$env:RSBuildEnableTests='true'
./build.ps1 -Configuration Release -Platform x64 -MaxCpuCount 4
./Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_SelectionShapeSmb -Configuration Release -Platform x64 -SkipBuild -TestRoot Z:/RedSalamander.Perf -AllowAlternateVolumeTestRoot
```

Use each exact suite/filter in `candidate-scope.json`; the managed cancellation entry additionally
uses `-SelfTestRepeat 3`. Viewer launch arguments, profile and environment provenance are recorded
in their qualification/trace files; all are the same `--no-activate` cases as the Debug packet.
Tooling summary names all nine input files and the excluded tag.

Each runtime directory preserves original native results/trace, governed aggregate, receipt and
final workspace snapshot. Selected metrics retain complete original JSONL line contents;
`metric-selection.json` binds every derivative to the full raw source digest and selection rule.
Complete bulk metrics and harness captures remain under
`Z:/RedSalamander.Perf/evidence/affected-master-release-2026-10-18/`.
`pre-staging-verification.json` verifies all selected results, source hashes and artifact bytes;
`tested-source-files.json` binds delivered source contents. A later staging/commit changes Git
metadata and does not constitute exact Resume reuse of these dirty-checkout receipts.

