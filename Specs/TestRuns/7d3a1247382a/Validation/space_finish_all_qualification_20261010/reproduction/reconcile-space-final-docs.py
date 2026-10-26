import argparse
import difflib
import hashlib
import json
import pathlib
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

parser = argparse.ArgumentParser()
parser.add_argument('--execute', action='store_true')
args = parser.parse_args()
repo = pathlib.Path(__file__).resolve().parents[2]
authority = pathlib.Path('D:/RedSalamander.Perf')
assert (authority / '.red-salamander-test-root').read_text().strip() == 'red-salamander.test-sandbox-root.v1'
evidence = authority / 'evidence/viewer-space-finish-all-20261009'
frozen = json.loads((evidence / 'queued-path-frozen-source.json').read_bytes())
matrix = json.loads((evidence / 'root-final-current-qualification-matrix.json').read_bytes())
assert matrix['qualifiedBuildSource'] == frozen['SnapshotId']
assert matrix['planStatus'] == 'ACTIVE' and not matrix['targetWaivers']
controlled = matrix['campaigns']['controlledRetained']['result']
assert controlled['status'] == 'qualified' and len(controlled['pairedContrasts']) == 6
ratio = controlled['candidateVsBaselineRatio']['pairedPeakPrivateDeltaRatioMedian']
input_result = matrix['currentInputLatency']['result']
assert input_result['mismatches'] == 0
input_samples = input_result['samples']
hover_ms = input_samples['hoverHandlerToNextPresentSubmission']['p95NearestRankUs'] / 1000
zoom_in_ms = input_samples['retainedControlWheelZoomInToTargetPresentSubmission']['p95NearestRankUs'] / 1000
zoom_out_ms = input_samples['retainedControlWheelZoomOutToTargetPresentSubmission']['p95NearestRankUs'] / 1000
input_run = input_result['provenance']['runId']
assert hover_ms <= 16 and zoom_in_ms <= 50 and zoom_out_ms <= 50

controlled_label = 'retained-reference-231508122z-399289c9'
ordinary_label = 'ordinary-current-224729245z-4221bb43-v2'
for label in [controlled_label, ordinary_label]:
    assert (repo / 'Specs/TestRuns/7d3a1247382a/Viewers' / (label + '-overview')).is_dir()
assert json.loads((evidence / 'root-ordinary-current-v2-replay.json').read_bytes())['mismatches'] == 0
assert json.loads((evidence / 'root-retained-reference-current-replay.json').read_bytes())['mismatches'] == 0

paths = ['Specs/Plans/WIP/ViewerSpace_ResponsiveTreemapAndAnalysis_2026-10-08.md',
         'Specs/Plans/WIP/README.md', 'Specs/Plugins/Plugins_ViewerSpace.md', 'Specs/NormativeConsistency.json']
before = {}
for path in paths:
    data = (repo / path).read_bytes()
    pin = next(item for item in frozen['WorkspaceSnapshot']['files'] if item['path'] == path)
    assert hashlib.sha256(data).hexdigest() == pin['worktree_sha256'], path
    before[path] = data.decode('utf-8')
after = dict(before)
current = f'''## Current implementation checklist — 2026-10-10

- [x] Implement Explorer and Analyst, nested folder/file tiles, folder headers and folded file corners, Rainbow/high-contrast cues, linked selection/filter/Inspector, responsive panels and in-flight scan totals. The durable behavior is in `Specs/Plugins/Plugins_ViewerSpace.md`.
- [x] Build the queued-path revision in test-enabled x64 Debug and Release with zero warnings/errors: source `9addaf53d4d88d618223a7f2d8554ac0b73f2472946c3a3e0875b1fd94ac9554`, Debug receipt `f721d4c3644b36016e3b81900978cc335f1a530d9eadc094f585683bc7f3638b`, Release receipt `37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd`.
- [x] Qualify the queued same-instance title/generation/no-geometry witness in Debug run `20261009T195007Z-29512-fd2830b98aa34a8787a044fd55288da5`. Both broad Commands runs recorded ViewerSpace 34 passed, 0 failed, 4 declared opt-in skips, including the 20K and 50K cases. The separate exact native activation/Escape Debug run `20261009T223523Z-45208-85e484e24c8249f78bb3e3a7221599bb` passed with zero remaining owned processes and exact saved foreground/keyboard focus. Seven noninteractive native ViewerSpace cases passed in each PrePush profile; Release broad interactive also reached and passed Space activation four times before later Web/Shell failures.
- [x] Review the 11 current Release Commands frames and eight proper lifecycle frames at original resolution, bound to the exact 144-DPI fixtures/builds. Refresh the three documentation PNGs and adjacent provenance from those immutable captures; exclude the extra image with limited provenance. Review records are in `Validation/space_finish_all_root_review_20261010`.
- [x] Complete current-source Debug and Release PrePush attempts and preserve their failures: Debug 2,127 passed / 17 failed / 261 skipped; Release 2,125 / 16 / 264, each 2,405 obligations. Both change/repository verdicts remain **FAILED**. Debug promoted 14 of 20 entries; Release 15 of 20. These are terminal attempts, not green repository qualification.
- [x] Align the baseline to exact candidate DxUi library `FB45C1FE673BBCF48F3CA2B6E104C7220B03990CAB375B215022EF92026EFA13` and receipt `935a4954a73b50d93f37a9df89f4f880aecf50ec4e5d55f7663ac0de13c8cf44`, source `3be7cc158093a805b7a0226e5a4bff7f9d33f792684b066ebe2b2125dea6984b`. Run the ordinary comparison: 14 passed runs / 28 cases, two excluded warmups and six adjacent ABBA contrasts per corpus. Root replayed 151 source/archive pairs, 59 generated digests and all 78,770 rows / 32,831,995 bytes. Corrected v2 archives passed the canonical validator on 240 files. Initial rejected byte-fragment naming remains preserved separately; measurements were not changed.
- [x] Run final cached-200 and original CandidateHeavy on Release receipt `37dae…`. Cached correctness passed all 200 exact-hit/current-presentation samples, but timing remains **MISS**: minimum 97.330 ms, median 117.765 ms, p95 145.603 ms, maximum 204.976 ms, 1/200 ≤100 ms. CandidateHeavy passed four equal 100,000-file/501-directory/3,001,292-byte roots, four cache hits and all 24 phases under unchanged guards. Root replayed 1,507 selected rows in each archive; these functional passes do not establish historical S0 memory gain.
- [x] Complete the controlled same-IViewer retained-reference comparison: 14 passed runs, two excluded warmups, six ABBA contrasts, four identical roots/four hits/25 phases per run. Median paired candidate/baseline sampled peak private-byte increment is **{ratio:.6f}**. This denominator is the controlled retained-reference lifecycle, not historical fresh-instance S0. Archive prefix: `Viewers/{controlled_label}`; root replay and exact source/runtime bindings accompany it.
- [x] Preserve exact current Release hover/navigation run `{input_run}`: 200 hover samples (p95 {hover_ms:.3f} ms), 200 Control-wheel zoom-in ({zoom_in_ms:.3f} ms) and 200 zoom-out ({zoom_out_ms:.3f} ms). All correctness assertions passed; root recomputed all 600 statistics and 4,215 retained raw rows. These meet the settled-fixture 16-ms/50-ms submission targets, with no baseline gain or scanning/selection/Up/Back/Forward claim. The scoped report is v1; NOT_EVALUATED is the canonical focused-Commands coverage classification, not an invented v2 verdict. Three disk-audit issues remain. Open-to-chrome first paint stays unqualified because its available metric is scan-start based.
- [ ] Resolve the unchanged absolute targets: small ordinary first meaningful map **MISS** (median 280.8665 ms, p95 302.094 ms, 0/6 ≤250 ms); cached first map **MISS** as above. Large ordinary first map passed its named corpus (median 186.462 ms, p95 197.794 ms, 6/6 ≤250 ms). With n=6, nearest-rank p95 is the maximum; broader tails remain unqualified.
- [ ] Respect §12's stop after two failed cached-target repair attempts. The user's pending choice between scoping graphics-device reuse and accepting a documented latency exception remains unanswered; no waiver is assumed. The device-slot feasibility proposal is not implemented. Historical ≤70%-of-S0 remains **UNMEASURED**, independently from the controlled result.
- [ ] Resolve separate owners' gates: Commands/Compare/FileOps/native/provider failures; Windows symbolic-link privilege; repository/account CI billing; native ARM64; live providers/devices; translation, broader DPI/real-device and human screen-reader review. Preserve exact outcomes, skipped reasons, resource guards and the Q9 desktop restoration contract; no environment or billing change is authorized.
- [ ] Move to Done only after §12's assertions, measurements and explicit dispositions are satisfied. The WIP remains **ACTIVE**. Final documentation/evidence edits retain the five tested code files byte-for-byte; the named receipts qualify their frozen source, not a new complete workspace snapshot.

Final root qualification matrix and replay records: `Specs/TestRuns/7d3a1247382a/Validation/space_finish_all_qualification_20261010`. Ordinary results: `Viewers/{ordinary_label}*`; current cached/heavy results: `Viewers/cached-200-224059615z-9e53a14b*` and `Viewers/candidate-heavy-224302601z-0b9ae69c*`. Debug/Release aggregate reports remain in their dated `Validation/space_*_prepush_reports_20261009T*` archives. `SOURCE_INPUT_CHANGED` is an impact-selection reason, not evidence of mutation during these frozen runs. Ordinary closed-after-quiet private-byte increment ratios were 1.29267 small and 1.29856 large; the peak reduction does not establish lower memory after Close.

## Historical 620d qualification and queued-path work — 2026-10-09

The following checklist preserves its then-current state. The 9add results above supersede only the same endpoint/profile; historical unbuilt/pending wording is not the present status.

'''
wip = paths[0]
start = after[wip].index('## Current implementation checklist')
end = after[wip].index('## Historical finish-all follow-through', start)
old = after[wip][start:end]
old = old[old.index('\n\n') + 2:]
after[wip] = after[wip][:start] + current + old + after[wip][end:]
after[wip] = after[wip].replace('## Previous implementation and qualification status — 2026-10-09', '## Historical f43 implementation and qualification status — 2026-10-09', 1)
table_start = after[wip].index('| §3.4 endpoint | Current status |')
table_end = after[wip].index('## Historical evidence ledger', table_start)
table = f'''| §3.4 endpoint | Current status |
|---|---|
| ≥25% matched ordinary scan improvement | Measured small/large median paired candidate/baseline ratios **0.042083 / 0.054444**; nearest-rank p95 **0.163159 / 0.141993**, n=6 each. Target met on these corpora only. Wide/skewed/provider and general ≤5% no-regression tails remain pending. |
| Small ordinary first meaningful map ≤250 ms | **MISS:** median 280.8665 ms, p95 302.094 ms, 0/6. Baseline meaningful-map timing unavailable. |
| Large ordinary first meaningful map ≤250 ms | **PASS for the named corpus:** median 186.462 ms, p95 197.794 ms, 6/6. Baseline meaningful-map timing unavailable. |
| Cached-model first meaningful map ≤100 ms | **MISS:** final candidate median 117.765 ms, p95 145.603 ms, 1/200. All 200 correctness/current-presentation witnesses passed. |
| Original CandidateHeavy functional lifecycle | Four equal roots, four hits, 24 phases passed under the unchanged 80-MiB/360-second/120-second-per-root guards. Functional completion only. |
| Historical memory-heavy private-byte increment ≤70% of S0 | **UNMEASURED:** no valid historical fresh-instance denominator. The controlled same-IViewer cohort does not replace S0. |
| Controlled same-IViewer retained-reference memory | Median paired sampled peak private-byte increment ratio **{ratio:.6f}**, six ABBA contrasts; controlled 70% comparison {'met' if ratio <= .70 else 'missed'}. Preserve its distinct lifecycle denominator and actual phase values. |
| Ordinary closed-after-quiet memory | Median paired private-byte increment ratios **1.29267 small / 1.29856 large**. No lower-memory-after-Close claim. |
| Settled hover-to-Present submission | **PASS on the named fixture:** p95 {hover_ms:.3f} ms, n=200. This is submission, not physical input-to-photon; scanning and selection remain unqualified. |
| Retained Control-wheel target-layout submission | **PASS on the named settled fixture:** zoom-in p95 {zoom_in_ms:.3f} ms, zoom-out {zoom_out_ms:.3f} ms, n=200 each. Up/Back/Forward, scanning and broader geometry remain unqualified. |
| Layout/analysis turn observations | Current filtered fixture: layout p95 0.355 ms/max 0.592 ms (n=403); analysis 0.231/0.336 ms (n=402), within 2-ms/4-ms budgets. This does not qualify the full deterministic stress corpus. |
| Open-to-chrome first paint ≤100 ms | **UNQUALIFIED endpoint gap:** available first-paint clock starts at scan start; Open setup and meaningful-map endpoints are distinct. |
| Other latency, layout/analysis, wide/skewed/provider, localization, platform and human accessibility tails | **PENDING** unless independently qualified in named archived evidence. |

The 25%, 70%, 250-ms and 100-ms targets are proposal-authored; all definitions and thresholds remain unchanged and none is waived. Documentation capture provenance is in `docs/res/viewer-space-captures.provenance.json`. The exact 9add build/run scope is recorded above; subsequent documentation/evidence edits do not make its receipts qualification of a new complete snapshot. All earlier snapshot-specific wording below remains historical.

'''
after[wip] = after[wip][:table_start] + after[wip][table_end:]
table_anchor = after[wip].index('## Historical 620d qualification')
after[wip] = after[wip][:table_anchor] + '### Current measured endpoints\n\n' + table + after[wip][table_anchor:]
question = '| Q9 | What time is acceptable for real keyboard/pointer/Narrator focus qualification? | The user authorized governed focus-taking work during this implementation session across broad Commands, File Operations, ViewerSqlite, ViewerPE and Monitor latency. Use the existing warning, owned-window desktop lease and cursor/foreground restoration. This authorizes execution, not a pass, manual screenshots or human accessibility qualification. | Respective local suite owners retain failed outcomes; accessibility reviewer owns human screen-reader assessment. |'
after[wip], count = re.subn(r'^\| Q9 \|.*$', question, after[wip], count=1, flags=re.MULTILINE)
assert count == 1
index = paths[1]
line = f'| ACTIVE | I31 | [ViewerSpace responsive treemap and linked analysis](ViewerSpace_ResponsiveTreemapAndAnalysis_2026-10-08.md) | Explorer/Analyst delivered; frozen 9add Debug/Release builds and scoped Space coverage passed. Both PrePush attempts remain FAILED (Debug 2127/17/261; Release 2125/16/264). Matched ordinary scan ratios 0.04208/0.05444; small first map MISS, large named corpus PASS. Cached-200 target MISS (1/200 ≤100 ms). Controlled same-IViewer memory ratio {ratio:.6f}, six ABBA contrasts; historical ≤70%-of-S0 UNMEASURED. Current settled hover/Control-wheel p95s {hover_ms:.3f}/{zoom_in_ms:.3f}/{zoom_out_ms:.3f} ms meet named-fixture targets. Corrected ordinary archives/root replay passed; current captures are source-bound. §12 latency disposition and named external/broad/human gates remain open; targets unchanged, WIP ACTIVE. |'
after[index], count = re.subn(r'^\| ACTIVE \| I31 \|.*$', line, after[index], count=1, flags=re.MULTILINE)
assert count == 1
plugin = paths[2]
old = 'A public `Close` releases the retained model and model-derived storage before returning.'
new = 'A public `Close` detaches the retained model and model-derived state from the active viewer before returning. Submitted deferred-disposal callbacks may retain retired storage until completion; `Close` does not wait for them.'
assert after[plugin].count(old) == 1
after[plugin] = after[plugin].replace(old, new)
after[plugin] = after[plugin].replace('Current x64 Debug governed PrePush fail: run `20261009T110605Z', 'Historical f43 x64 Debug governed PrePush fail: run `20261009T110605Z', 1)
after[plugin] = after[plugin].replace('Targeted same-instance qualification: source `620d85a', 'Historical targeted same-instance qualification: source `620d85a', 1)
qualification = f'''\n### Frozen-source qualification follow-up — 2026-10-10

The queued-path revision at source `9addaf53d4d88d618223a7f2d8554ac0b73f2472946c3a3e0875b1fd94ac9554` built in test-enabled x64 Debug and Release with zero warnings/errors; receipts are `f721d4c3644b36016e3b81900978cc335f1a530d9eadc094f585683bc7f3638b` and `37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd`. Both broad Commands runs recorded Space 34 passed/0 failed/4 opt-in skipped, including the 20K/50K cases and queued same-instance reopen. Seven noninteractive native Space cases passed in each profile. Exact Debug activation/Escape run `20261009T223523Z-45208-85e484e24c8249f78bb3e3a7221599bb` passed with exact saved focus and zero remaining owned processes. Release broad interactive passed Space activation four times before later Web/Shell failures. Current 11-frame Commands and eight proper lifecycle captures are bound to their 144-DPI fixture, actual binary hashes and immutable receipt; adjacent documentation provenance retains earlier captures separately.

Current PrePush attempts remain **FAILED**: Debug `20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f`, 2,127 passed/17 failed/261 skipped; Release `20261009T213050Z-43264-4bde93ea855e44c885a3277a68e89038`, 2,125/16/264, each 2,405 obligations. Both change/repository verdicts are FAILED. Three sandbox-audit issues remain in scoped runs. Preserve Commands/Compare/FileOps/provider/native failures and exact skip reasons. General S3 Debug assertions failed in Debug; their Release gate was skipped, so the Release PluginContract pass does not replace those assertions. `SOURCE_INPUT_CHANGED` is impact-selection metadata, not proof of mutation during the frozen runs.

Matched ordinary campaign `20261009T224729245Z-4221bb43` used baseline receipt `935a4954a73b50d93f37a9df89f4f880aecf50ec4e5d55f7663ac0de13c8cf44`, source `3be7cc158093a805b7a0226e5a4bff7f9d33f792684b066ebe2b2125dea6984b`, and the candidate Release receipt with identical DxUi `FB45C1FE673BBCF48F3CA2B6E104C7220B03990CAB375B215022EF92026EFA13`. Fourteen runs/28 cases passed, with two excluded warmups and six ABBA contrasts per corpus. Median paired scan ratios were 0.042083 small / 0.054444 large. Candidate first meaningful map: small **MISS**, median 280.8665 ms/p95 302.094 ms/0 of 6 ≤250 ms; large **PASS for this corpus**, median 186.462 ms/p95 197.794 ms/6 of 6. Baseline meaningful-map timing was unavailable; n=6 nearest-rank p95 is the maximum. Large sampled peak private-byte increment ratio was 0.618203; closed-after-quiet ratios were 1.29267 small / 1.29856 large, so this does not establish lower memory after Close. The corrected `{ordinary_label}*` archives passed validation on 240 files; root reconstructed all 78,770 raw rows and replayed 151 source/archive pairs plus 59 generated hashes with no mismatch. Earlier byte-fragment naming failure is preserved as rejected curation, with unchanged measurements.

Final cached-200 run `20261009T224102Z-43236-26566751591f4c62bbb2370dea3e1e75` passed all 200 exact-hit/current-presented samples but remains **MISS** at the unchanged 100-ms target: median 117.765 ms, p95 145.603 ms, 1/200. Original CandidateHeavy `20261009T224305Z-42696-6a56004e8deb44eb938f09454d6503bd` passed four equal roots/four hits/24 phases under unchanged guards. Controlled retained-reference campaign `20261009T231508122Z-399289c9` passed 14 runs, two excluded warmups and six ABBA contrasts with four equal roots/four hits/25 phases per run. Its median paired sampled peak private-byte increment ratio was {ratio:.6f}; the denominator is the explicitly controlled same-IViewer lifecycle. Historical fresh-instance ≤70%-of-S0 remains **UNMEASURED**. These cohorts do not qualify broad wide/skewed/provider tails or general speed/memory claims.

Current filtered Release hover/navigation run `{input_run}` passed its exact case with 200 hover and 200 Control-wheel observations per direction. Root replayed all 600 sample statistics and 4,215 byte-exact selected raw rows. Settled submission p95s were {hover_ms:.3f} ms hover, {zoom_in_ms:.3f} ms zoom-in and {zoom_out_ms:.3f} ms zoom-out; the respective 16-ms/50-ms targets are met for this fixture. Layout-turn p95/max were 0.355/0.592 ms (403 rows), analysis 0.231/0.336 ms (402 rows); these observations do not qualify the full stress corpus. The actual scoped report schema is v1, and NOT_EVALUATED derives from the canonical focused-Commands coverage contract. Its three disk-audit issues remain recorded. No baseline gain, input-to-photon, scanning/selection/Up/Back/Forward or 100-ms Open-to-chrome qualification is inferred; available chrome first-paint instrumentation is scan-start based.

Authoritative behavior and result qualification are separate. Current evidence/root replay is retained under `Specs/TestRuns/7d3a1247382a/Validation/space_finish_all_qualification_20261010` and the active ViewerSpace WIP; current screenshot review is under `Validation/space_finish_all_root_review_20261010`. Final documentation/evidence edits preserve the five tested code files byte-for-byte; the named receipts qualify their frozen source, not a new complete workspace snapshot. §12's cached-latency stop and user disposition remain pending. Hosted CI billing, Windows symbolic-link privilege, native ARM64, live providers/devices, broader DPI/translation and human accessibility retain named separate owners. No target is waived and the plan remains ACTIVE.
'''
after[plugin] += qualification
norm = paths[3]
document = json.loads(after[norm])
entry = next(item for item in document['documents'] if item['path'] == plugin)
old_note = entry['reviewerNote']
new_note = old_note + ' Targeted 2026-10-10 frozen-source follow-up: Close detaches active model state while submitted disposal callbacks may retain storage; same-instance Open starts with an empty model before creation callbacks. Current 9add qualification, controlled retained-reference measurements, exact 144-DPI captures and preserved FAILED/MISS/UNMEASURED verdicts are recorded in the dated Testing Contract follow-up and active WIP. This targeted review does not restamp whole-document metadata or qualify unrun platforms, providers, locales, hosted CI or human accessibility.'
old_literal = json.dumps(old_note, ensure_ascii=False)
assert after[norm].count(old_literal) == 1
after[norm] = after[norm].replace(old_literal, json.dumps(new_note, ensure_ascii=False), 1)
assert json.loads(after[norm])['reviewedAtCommit'] == document['reviewedAtCommit']
for path in paths:
    assert after[path] != before[path]
    print(''.join(difflib.unified_diff(before[path].splitlines(keepends=True), after[path].splitlines(keepends=True), fromfile=path, tofile=path)))
if args.execute:
    for path in paths:
        (repo / path).write_text(after[path], encoding='utf-8', newline='')
    record = {'reviewer': 'root', 'qualifiedCodeSource': frozen['SnapshotId'], 'scope': 'Four targeted documentation updates; code remains frozen. No waiver or Done move.',
              'files': [{'path': path, 'beforeSha256': hashlib.sha256(before[path].encode()).hexdigest(),
                         'afterSha256': hashlib.sha256(after[path].encode()).hexdigest()} for path in paths]}
    with (evidence / 'root-final-documentation-reconciliation.json').open('x', encoding='utf-8') as stream:
        json.dump(record, stream, indent=2)
        stream.write('\n')
    print('Applied four targeted documentation updates.')
else:
    print('PREVIEW_ONLY: no tracked files written.')
