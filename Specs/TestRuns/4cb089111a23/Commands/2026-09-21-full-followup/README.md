# Fresh Full failure and focused follow-up

Source base `e5f735413e7cd932bc342bdd7c79be859199941a`, qualified DxUi pin
`78b3de389a189c7f86f611787e0489fb6d474218`. Full run
`20260921T190208Z-107796-357e522a509840379dc9046cc0ed6aa1` exited 1:
Commands **912 passed, 2 failed, 2 optional skips**; FileOps **192 passed,
1 failed, 20 capability skips**. Other Full entries promoted. Raw copied
case results preserve all skips and failure reasons. Original complete evidence
remains under `C:/RedSalamander.Perf/evidence/runs/<run-id>/`.

The throughput external-reload test required a painted summary for a card that
could be outside the viewport. Its revised fixture holds a 720×260-DIP client
in normal sizing state, verifies remembered expansion separately, scrolls the
appended card into view and pumps for actual paint. No production behavior or
settings-save assertion changes. Test-enabled full Debug receipt:
`20153c27c17fe93b4c6892ce3f7b6efa0418058ff871eb8f55f7db2c1689a8f9`.

- Throughput disclosure/reload/restart: **9/9**, three repetitions, runner
  `20260921T204000Z-97332-6d4296aa2e68446bb32aacec8105d42b`.
- Selection visibility alone: **3/3**, runner
  `20260921T204728Z-115864-1b136283476e4ab7bd1f32083248a231`.
- Surrounding `folder_view_` selection/focus cluster: **30/30**, runner
  `20260921T205003Z-112284-31432e8b5a2c474aaf6d4c2ee366174b`.
- Unchanged C/D independent volumes: **9/9** including setup/cleanup, runner
  `20260921T205155Z-109384-8d2a73af8bb54a9a953511a1b018599c`.

All focused runs have zero skips. The first two retain runner logs only: their
raw scratch was rotated before copying. Cluster and volume results were copied
before the next runner. An earlier volume attempt lacked the already-authorized
`-AllowAlternateVolumeTestRoot` flag and was blocked; it is not timing evidence.
Its failed log is retained externally as `volumes-baseline.log`.

Unchanged isolated C/D/concurrent timings (microseconds):
1,265,000 / 1,203,000 / 1,219,000;
1,297,000 / 1,172,000 / 1,250,000;
1,312,000 / 1,187,000 / 1,312,000.
The earlier Full concurrent duration was 2,453,000 versus a 1,709,100 gate.
This difference remains unresolved; isolated passes do not establish Full order
cleanliness. The benchmark currently includes unequal fixture work before task
admission. The follow-up corrects that boundary for all three samples, preserving
the 1.35 ratio, overlap, data checks and absolute timeout. Its candidate results
are retained below. No product selection repair or Full pass is claimed.

Candidate Debug receipt `b873dc708ebc10e3c465344476a5a3a78d5e7b7074b128ab0c57b98044c4b00e`:
the 13 exact Commands cases repeated three times pass 39/39; the discovery family
passes 15/15. Canonical Commands run
`20260921T214522Z-51540-4616ba19df6d411a9c0bc5f3d6475aa6` passes 914 cases,
zero failures and two explicit optional skips, including both earlier failed cases.

Canonical FileOps run `20260921T221901Z-30224-48c33229fb41493185701baba216f5f9`
aborted inside the provider capability matrix. Its six child results are not parent
coverage; the runner correctly rejects the incomplete result set. Matching-symbol
dump analysis found roughly 897 KiB in the matrix's Debug frame with 41,808-byte Task
objects. Nested managed-move I/O exhausted the stack; crash-path formatting then
faulted too. The full dump and symbols remain external (see dump reference/hash).
The candidate moves 16 test-owned Task objects to scoped unique ownership without
changing assertions, timeouts or production task allocation. The full Debug rebuild
passes with zero warnings/errors, receipt
`f12674d8b70c03a02932fdab6a499da49ec0946d5003209501d4aa58103b9805`.
The exact matrix passes three consecutive executions (9/9 with setup/cleanup, zero
skips), run `20260921T224105Z-88448-33a72a4288144f75a37b505de0d9cfe4`.
Canonical FileOps run `20260921T224235Z-106584-346d000470cd46038890260d776a4865`
then passes all 193 cases with zero failures and 20 explicit capability skips in
1,351,217 ms. The runner returns 1 because its disk audit finds 260 historical
directories, so this is a functional suite pass but not an overall qualification
pass. Both raw receipts are retained. Spec inventory reports zero blocking findings.

On September 23, the supported root initializer created `E:/RedSalamander.Perf`;
its audit has zero findings. No existing C:/D: run or evidence directory was moved,
deleted or allowlisted. Future final qualification must execute on the clean root;
the new root does not change any earlier run's outcome.

`volumes-baseline-metrics.jsonl` is a claim-bearing subset in original order.
Complete external inputs: `C:/RedSalamander.Perf/evidence/i26-ui/full-followup-20260921/`.
The runner's 260 disk-audit findings remain explicit. This packet does not close
Full/Release, native DPI/AT, resource qualification or I26.
