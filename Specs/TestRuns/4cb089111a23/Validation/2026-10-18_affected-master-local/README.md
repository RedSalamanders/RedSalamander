# Affected local validation after the 2026-10-18 master merge

Requested scope: run tests affected by the merged modifications locally while hosted CI cannot start,
repair failures, update the owning contracts and preserve validation evidence. This is selected x64
Debug validation, not Fresh Full, native ARM64, ASan or a clean-sandbox acceptance claim.

## Candidate and selection

The delivery merge is `9a4d243416fa503899c2f44e4242a1276db1cef9`, compared with pre-merge master
`265147ee62c296b1a47ed44efb4d8edb58ad60cf`. `initial-scope.json` maps the merged production paths
to explicit cases. The generic explanation retains its broader conservative recommendation in
`explain-plan.json`; that recommendation is not a passing result or permission to omit affected cases.

Selected runtime surfaces are natural display sorting and retained pane state; Find, Batch Rename,
autocomplete and breadcrumbs; decision-mirror keyboard/focus policy; discovery scope accounting and
cancel/queue transitions; shared bridge scheduling and destination budgets; actual disk-to-SMB
Copy/Move; and first-frame completed-card geometry. Three native ViewerSpace cases cover the other
consumer of the shared natural comparison. Tooling selection follows the changed plan/spec modules'
declared tests plus inventory/governance, workflow-pin policy and reviewed scope identity.

No Compare, unrelated viewer/device runtime, installer, monitor, build-toolchain integration or entire
repository suite is claimed from this explicit selection. The legacy 350-file/two-root SMB cost gate
and independent qualification plans keep their original status.

## Reproduced failures and repairs

1. Initial `FileOpsFamily_DiscoveryScope` run
   `20261018T180555Z-81544-e237aa74d67549678b9a43aa45f39193` failed its first leaf case: it expected
   closing discovery to change `Copy — payload.bin` into a denominator-bearing heading. The current
   UI contract retains operation/subject identity and places progress in a separate primary amount
   row. The repaired Copy/Move and empty/verified-leaf checks assert the unchanged heading and read
   the actual popup's painted amount, with exact byte or file denominator, at the held pre-payload
   checkpoint. Producer totals, determinate/compact progress, final byte counts and source outcomes remain
   independently asserted. The failed run has two setup/cleanup passes, one failure and 24 cases not
   reached; none of those skipped cases is promoted to coverage.
2. The historical-profile baseline passed three tests and failed the existing Gate0 SHA-256 guard.
   Dependabot had changed three artifact-action references in the exact-byte historical workflow.
   Restoring the file from `3e84e87b7` restores its existing sealed digest
   `e6f15b6fef9c0d367e12dc7f43cb769106dc982934658f501120728599f32d1c`; no seal or historical test
   is rebaselined. Dependabot now excludes only that exact path. Current workflow artifact actions
   still require current reviewed pins; the historical exception requires the entire sealed file hash.
   The focused five-test guard and the final 421-check affected tooling selection pass after repair.
3. The bridge/connection family passed its native assertions, but its pipeline metric baseline
   included 4,096 delayed seed writes: 129.234 seconds versus a 0.297-second candidate. That pair is
   **not qualified** for an improvement claim. Setup now uses zero simulated stream latency, both
   timers start immediately before admission, and one 128 MiB file isolates within-file overlap
   while retaining the original total payload and 80-percent threshold. Exact completed bytes are
   required in both modes. Its corrected pair records 157 ms setup outside the timers, 1.422 seconds
   disabled and 0.922 seconds enabled. The old and revised corpora are different; their timings
   must not be compared as a before/after product speedup. See `pipeline-observation-review.json`.

GitHub documents the per-ecosystem file exclusion in its
[Dependabot options reference](https://docs.github.com/en/code-security/reference/supply-chain-security/dependabot-options-reference#exclude-paths).

## Candidate identity

Test-enabled x64 Debug full-solution rebuild: zero warnings/errors, receipt
`b1ac26bc09c7963d9192ddb4a573e6ea0198650800a31f439585d3190bdf0f4d`.
The tested dirty source snapshot is
`06bd32c8dac88276328262dfe952060ea2dda9a13cfc4afebb4f30d79d569933`, based on merge `9a4d24341`.
`candidate-source.json` records every repair input; `candidate-scope.json` records exact selections.
Every governed run verifies the solution receipt and its artifact bytes; the orchestration also
compares the canonical workspace snapshot after every selection. A later commit changes Git metadata
and is not represented as exact Resume reuse of these results.

The subsequent pipeline-only test repair and its owning performance-spec paragraph bind final
receipt `d26c1a3b04bb9469c36e907e5315730dc7793d4979bd07aa11ae78704e79ba66`, snapshot
`aeae66efb25b50ba6b8bdaf4a1119b54fee400a5477bb4398f05b7f43c4cf443`. That rebuild also has zero
warnings/errors. `pipeline-source-delta.json` independently records the two changed paths and zero
production C++ changes. The exact repaired pipeline case, alternate-volume retry, three viewer cases
and affected tooling use this final identity. Earlier selections retain their own immutable identity;
they are not represented as fresh executions against the later executable.

## Completed selected validation

- Repaired discovery scope group: 27 passed, zero failed/skipped.
- Candidate Commands selection: 11 passed, zero failed/skipped, including the 5,000-row sort scenario
  and real decision-mirror focus test through the existing warning and desktop lease.
- Candidate decision-mirror group: nine passed, zero failed/skipped.
- Candidate discovery cancel/queue group: 11 passed, zero failed/skipped.
- Bridge/connection group: nine native passes; its original pipeline timing pair remains unqualified
  and is replaced by the corrected exact-case pass above.
- Native SMB budget and managed Move staging: three passes each, zero failed/skipped.
- Destination-storage case: its first invocation skipped the case because the alternate-volume
  switch was absent. The retry explicitly enables the already marked, configured
  `D:/RedSalamander.Perf` sandbox and passes all three rows with zero failed/skipped.
- Local and real SMB selection-shape cases: three passes each, zero failed/skipped. Independent
  metric checks require all six shape rows per case, six four-file SMB rows, all five managed cleanup
  controls and the nested Move. The primary and alternate sandboxes retain their ownership markers.
- Completed-card remediation: three passes, zero failed/skipped, with all four collapse control rows.
- Repaired pipeline exact case: three passes, zero failed/skipped; exact payload and unchanged
  improvement threshold pass with fixture seeding excluded.
- ViewerSpace: all three selected native cases return zero with no activation, stable source identity
  and receipt-verified artifacts.
- Tooling: 421 passed, zero failed/skipped across the nine explicitly selected files. The wrapper's
  post-test snapshot command needed its canonical module reimported after Pester unloaded exports;
  the recovered barrier matches the archived pre-tooling source and every attested artifact. The
  assertions were not rerun or promoted as formal Resume evidence. A separate initial viewer helper
  omitted the receipt schema parameter and executed zero native cases; its pre-execution record is
  retained beside the three successful launches.
- Final static checks: 34 changed native files pass the pinned formatter, whitespace checks pass,
  all 197 tool files are classified, and specification inventory has zero blocking findings.

`summary.json` preserves the initial attempts. `candidate-summary.json` records each completed
candidate selection, including original run IDs and receipt/source identity. Each run directory keeps
native results, trace, governed aggregate and final workspace snapshot unchanged. Selected JSONL files
contain original complete matching lines; `metric-selection.json` binds each derivative to its full
raw source and exact selection. Complete bulk metrics remain in the marked sandbox under
`Z:/RedSalamander.Perf/evidence/affected-merged-master-2026-10-18/`.
`pipeline-summary.json` records the final pipeline and authorized alternate-volume retry; viewer
and tooling subdirectories keep their separate scope, source and artifact barriers. Explicit archive
validation is performed before staging this packet.

The runner still reports 41 preexisting unowned historical run directories; they were preserved.
Hosted CI on the delivery merge stopped before any steps because of account billing/spending limits.
No hosted result, historical performance observation or prior Full verdict is rewritten by this packet.
