from pathlib import Path
import json, hashlib, shutil, subprocess, re, statistics, csv
from collections import Counter, defaultdict
from datetime import datetime, timezone

REPO = Path(r'Z:\src\.worktrees\RS-audit20261113')
SCRATCH = Path(__file__).parent
PACKET = SCRATCH / 'packet-debug-v3'
SEALED = Path(r'Z:\RedSalamander.Perf\evidence\completed-runs')
ARCHIVES = Path(r'Z:\RedSalamander.Perf\evidence\run-archive-records')
HEAD = 'e83dd755f44cef4f720d3f439f5fbd06383a569a'
SNAPSHOT = 'bf6aed7cb9a8479f998dbde44717845600a8735933c4387f518d531e8100b04d'
RECEIPT_ID = '2a421a6b4d303b142f68a56faf7b070fdeff32953a4803884fba1ecd0e105135'
BASE = '78a8a75c46b784c52c5aeb82c0a3bf0773e61f6e'
BASE_HASH = '88662a7c5611143f0778969a56e7b5427298ff78fef4c43c718dab9f7e60b75a'
RUNS = {
    'Review': ('20261113T165013Z-78108-95535f9d947843e7b9305611372ac711', 'FileOps_ReviewRemediationContracts', 'a3aaae65b2aa08caacfcc074f2420f2c854053f76a74f268b16f7368edb53cba'),
    'Identity': ('20261113T165333Z-71548-2273ce8cd6be4862af7db946309da83c', 'FileOps_IdentityInterlockWaiting', 'ad68d11a6cd45008af9730637ccd40b6f778ae1e84b32f895ccc51afdeb7a6a4'),
}
METRICS = {
    'FileOps.Transfer.InsideSourceProbeUs', 'FileOps.SelfTest.LocalAncestryProbe',
    'FileOps.SelfTest.LocalCrossRootDistinctness', 'FileOps.Conflict.LocalDistinctnessUnproved',
    'FileOps.Power.ConflictWaitCalls', 'FileOps.Conflict.WaitCalls',
    'FileOps.SelfTest.QueueStartOrder.Us', 'FileOps.Queue.Selector.BaselineBatchUs',
    'FileOps.Queue.Selector.CandidateBatchUs', 'FileOps.Queue.Selector.CollectionCapacityBytes',
}

def sha(data): return hashlib.sha256(data).hexdigest()
def read_json(path): return json.loads(path.read_bytes())
def write_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + '\n', encoding='utf-8', newline='\n')
def utc(value): return datetime.fromisoformat(value.replace('Z', '+00:00'))
def git(*args): return subprocess.check_output(['git', '--no-optional-locks', *args], cwd=REPO)

assert not PACKET.exists(), 'Packet is create-only.'
PACKET.mkdir()
copies = []
def copy_exact(source, relative):
    data = source.read_bytes()
    destination = PACKET / relative
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(data)
    assert destination.read_bytes() == data
    copies.append({'source': str(source), 'packet_path': str(relative).replace('\\', '/'), 'bytes': len(data), 'sha256': sha(data)})
    return data

receipt_path = REPO / f'.build/x64/Debug/receipts/{RECEIPT_ID}.json'
receipt_bytes = copy_exact(receipt_path, Path('provenance/build-receipt.json'))
receipt = json.loads(receipt_bytes)
assert receipt['receipt_id'] == RECEIPT_ID and receipt['git_head'] == HEAD and receipt['source_snapshot_id'] == SNAPSHOT
assert receipt['configuration'] == 'Debug' and receipt['platform'] == 'x64' and receipt['tests_enabled'] and receipt['exit_code'] == 0
assert git('rev-parse', 'HEAD').decode().strip() == HEAD
assert not git('status', '--porcelain').strip(), 'Source was expected frozen/clean.'

build_context = SEALED/'20261113T164233Z-54404-7c6a1aaf2f3446578555674b91a20dd7'
sealed_receipt = build_context/'scratch/audit-native/full-debug-build/build-receipt.json'
assert sealed_receipt.read_bytes() == receipt_bytes
copy_exact(sealed_receipt, Path('provenance/sealed-build-receipt.json'))
copy_exact(build_context/'scratch/audit-native/full-debug-build/build-result.json', Path('provenance/sealed-build-result.json'))
selected_output_checks = []
for relative in ('.build/x64/Debug/RedSalamander.exe', '.build/x64/Debug/Plugins/FileSystem.dll', '.build/x64/Debug/Common.dll'):
    output = next(item for item in receipt['outputs'] if item['relative_path'] == relative)
    selected_output_checks.append({'relative_path': relative, 'sha256': output['sha256'], 'size_bytes': output['size_bytes'],
                                   'verification': 'historical immutable receipt output identity; current profile is being re-attested and is deliberately not used as historical proof'})

bindings = {}
selected_rows = {}
for label, (run_id, case_filter, expected_raw_hash) in RUNS.items():
    root = SEALED/run_id
    last = root/'artifacts/selftest/last_run'
    out = Path('Debug')/label
    for relative in ('results.json', 'trace.txt', 'run-all-tests-results.json', 'run-all-tests-case-history.jsonl',
                     'workspace-after-selftest.file-operations.json', 'workspace-final.json', 'fileops/results.json', 'fileops/trace.txt'):
        copy_exact(last/relative, out/relative)
    previous = root/'artifacts/selftest/previous_run'
    for name in ('entry-fingerprints.json', 'validation-plan.json', 'validation-plan-decision.json'):
        copy_exact(previous/name, out/'prelaunch'/name)
    fingerprints = read_json(previous/'entry-fingerprints.json')
    plan = read_json(previous/'validation-plan.json')
    decision = read_json(previous/'validation-plan-decision.json')
    assert fingerprints['build_receipt_id'] == RECEIPT_ID and fingerprints['snapshot_id'] == SNAPSHOT
    assert fingerprints['plan_digest'] == plan['plan_digest'] == decision['plan_digest']
    assert decision['snapshot_id'] == SNAPSHOT and plan['coverage_scope'] == 'filtered' and fingerprints['artifact_epoch'] == 0
    outer = read_json(last/'run-all-tests-results.json')
    native = read_json(last/'fileops/results.json')
    final = read_json(last/'workspace-final.json')
    after = read_json(last/'workspace-after-selftest.file-operations.json')
    assert outer['run_id'] == run_id and outer['case_filter'] == case_filter and outer['configuration'] == 'Debug' and outer['platform'] == 'x64'
    assert (native['passed'], native['failed'], native['skipped']) == (3, 0, 0)
    assert [item['name'] for item in native['cases']] == ['Setup', case_filter, 'Cleanup_RestorePluginConfig']
    assert all(item['status'] == 'passed' for item in native['cases'])
    assert final['snapshot_id'] == after['snapshot_id'] == receipt['source_snapshot_id']
    assert final['git_head'] == after['git_head'] == receipt['git_head']
    assert final['repo_identity'] == receipt['repo_identity']
    native_suite = next(item for item in outer['suites'] if item['suite'] == 'FileOperations')
    assert native_suite['exit_code'] == 0 and native_suite['classification'] == 'PASSED'

    process_files = list((root/'artifacts/native-process-exits').glob('*.json'))
    assert len(process_files) == 1
    exit_data = copy_exact(process_files[0], out/'native-process-exit.json')
    process_exit = json.loads(exit_data)
    assert process_exit['exit_code'] == 0 and process_exit['executable'] == outer['executable']
    records = ARCHIVES/run_id
    intent_data = copy_exact(records/'intent.json', out/'archive-intent.json')
    complete_data = copy_exact(records/'complete.json', out/'archive-complete.json')
    intent, complete = json.loads(intent_data), json.loads(complete_data)
    assert complete['intent_sha256'].lower() == sha(intent_data)
    assert intent['run_id'] == run_id and intent['disposition'] == 'completed'
    assert complete['destination'] == str(root)
    for name in ('entry-fingerprints.json', 'validation-plan.json', 'validation-plan-decision.json'):
        member = next(item for item in intent['inventory'] if item['path'] == 'artifacts\\selftest\\previous_run\\'+name)
        assert member['sha256'].lower() == sha((previous/name).read_bytes())

    raw_path = last/'perf/perf_metrics.jsonl'
    raw = raw_path.read_bytes()
    assert sha(raw) == expected_raw_hash
    member = next(item for item in intent['inventory'] if item['path'] == r'artifacts\selftest\last_run\perf\perf_metrics.jsonl')
    assert member['sha256'].lower() == expected_raw_hash and member['length'] == len(raw)
    lines = raw.splitlines(keepends=True)
    parsed = [(number, line, json.loads(line)) for number, line in enumerate(lines, 1) if line.strip()]
    assert all(row['build'] == 'Debug' and row['machineHash'] == '4cb089111a23' for _, _, row in parsed)
    assert {row['process'] for _, _, row in parsed} == {process_exit['process_id']}
    assert {row.get('commit', '') for _, _, row in parsed} == {''}, 'Do not replace blank native commit fields.'
    minimum = min(utc(row['timestamp']) for _, _, row in parsed)
    maximum = max(utc(row['timestamp']) for _, _, row in parsed)
    assert utc(outer['run_started_utc']) <= minimum <= maximum <= utc(outer['run_ended_utc'])
    assert maximum <= utc(process_exit['observed_utc'])
    chosen = [(number, line, row) for number, line, row in parsed if row['metric'] in METRICS]
    selected_rows[label] = [row for _, _, row in chosen]
    excerpt = b''.join(line for _, line, _ in chosen)
    excerpt_path = PACKET/out/'perf/perf_metrics.jsonl'
    excerpt_path.parent.mkdir(parents=True, exist_ok=True)
    excerpt_path.write_bytes(excerpt)
    write_json(PACKET/out/'perf/raw-row-index.json', {
        'schema': 'red-salamander.native-component-row-selection.v1',
        'source_path': str(raw_path), 'source_sha256': expected_raw_hash, 'source_bytes': len(raw),
        'source_physical_lines': len(lines), 'source_nonblank_rows': len(parsed),
        'selection': 'exact metric-name allowlist; raw original line bytes and original order; no metadata rewriting',
        'metric_allowlist': sorted(METRICS), 'excerpt_sha256': sha(excerpt), 'excerpt_bytes': len(excerpt),
        'excerpt_rows': len(chosen),
        'rows': [{'excerpt_line': n, 'source_line': number, 'line_bytes': len(line), 'line_sha256': sha(line), 'metric': row['metric']}
                 for n, (number, line, row) in enumerate(chosen, 1)],
    })
    # Preserve the complete original audit object, including exact path/reason list.
    audit = outer['test_sandbox_audit']
    write_json(PACKET/out/'sandbox-audit.json', audit)
    categories = dict(sorted(Counter(item['category'] for item in audit['issues']).items()))
    bindings[label] = {
        'run_id': run_id, 'case_filter': case_filter, 'sealed_root': str(root),
        'native_counts': {'passed': native['passed'], 'failed': native['failed'], 'skipped': native['skipped']},
        'native_process_exit_code': process_exit['exit_code'], 'native_suite_classification': native_suite['classification'],
        'outer_wrapper_exit_code': outer['exit_code'], 'whole_repository_qualification': 'NOT_EVALUATED; exact filtered v1 diagnostic evidence',
        'sandbox_audit': {'is_clean': audit['is_clean'], 'issue_count': audit['issue_count'], 'category_counts': categories,
                          'contract': 'Testing_SelfTests.md:644 calls this aggregate audit non-blocking; keep residue visible, never waive or relabel root clean',
                          'ownership': 'Unproven for unrelated listed residue. No issue is a current selected-run escape/reparse finding; no cleanup performed.'},
        'source_binding': {'git_head_from_outer_workspace': final['git_head'], 'receipt_git_head': receipt['git_head'],
                           'after_and_final_snapshot_id': final['snapshot_id'], 'receipt_source_snapshot_id': receipt['source_snapshot_id'],
                           'repo_identity': final['repo_identity'], 'receipt_id': RECEIPT_ID,
                           'prelaunch_entry_fingerprints_sha256': sha((previous/'entry-fingerprints.json').read_bytes()),
                           'prelaunch_fingerprints_digest': fingerprints['fingerprints_digest'],
                           'prelaunch_plan_digest': plan['plan_digest'], 'prelaunch_artifact_epoch': fingerprints['artifact_epoch'],
                           'native_metric_commit_field': '', 'native_commit_was_not_invented': True,
                           'join': 'sealed prelaunch fingerprints explicitly bind receipt ID, source snapshot, artifact epoch and filtered plan; plan/decision/fingerprints match and archive inventory binds their bytes; same exact executable path in outer aggregate and owned exit record; one native PID throughout hashed raw metrics; metric timestamps contained by outer interval and precede owned exit; matching after/final workspace snapshot and receipt; archive transaction binds raw bytes'},
        'process_id': process_exit['process_id'], 'outer_started_utc': outer['run_started_utc'], 'outer_ended_utc': outer['run_ended_utc'],
        'native_metric_first_utc': minimum.isoformat(), 'native_metric_last_utc': maximum.isoformat(),
        'native_metric_run_ids': sorted({row['runId'] for _, _, row in parsed}),
        'raw_metrics_sha256': expected_raw_hash, 'raw_metrics_bytes': len(raw), 'raw_metrics_rows': len(parsed),
        'excerpt_metrics_rows': len(chosen), 'metric_counts': dict(sorted(Counter(row['metric'] for _, _, row in chosen).items())),
        'archive_intent_sha256': sha(intent_data), 'archive_complete_sha256': sha(complete_data),
    }

# Independent count/fixture/observer checks over the exact selected rows.
review = selected_rows['Review']
guards = [row for row in review if row['metric'] == 'FileOps.SelfTest.LocalCrossRootDistinctness']
assert len(guards) == 3 and {row['detail'] for row in guards} == {'same-entry', 'different-entry', 'identityless-replace'}
assert all((row['value0'], row['value1'], row['hr']) == (1, 0, 0) for row in guards)
power = [row for row in review if row['metric'] == 'FileOps.Power.ConflictWaitCalls']
assert len(power) == 6 and {row['value0'] for row in power} == set(range(6))
power_names = ['Preanswered', 'Precancelled', 'QuietCancel', 'QuietResume', 'StopOnly', 'Spurious']
power_maximum = [2, 2, 2, 6, 2, 3]
assert all(row['hr'] == 0 and row['value1'] <= power_maximum[row['value0']] for row in power)
assert all(row['hr'] == 0 for row in review if row['metric'] == 'FileOps.SelfTest.LocalAncestryProbe')

queue = [row for row in selected_rows['Identity'] if row['metric'].startswith('FileOps.Queue.Selector.')]
assert len(queue) == 144
pairs = []
for offset in range(0, len(queue), 3):
    # Preserve and verify actual emission order. Detail repeats across fixture sizes;
    # capacity value0 is bytes, so duration value0 supplies that triplet's task count.
    b, c, cap = queue[offset:offset+3]
    assert [row['metric'] for row in (b,c,cap)] == ['FileOps.Queue.Selector.BaselineBatchUs', 'FileOps.Queue.Selector.CandidateBatchUs', 'FileOps.Queue.Selector.CollectionCapacityBytes']
    detail = b['detail']
    assert c['detail'] == cap['detail'] == detail and detail.split(';')[0] == 'queue_head_selector_v1'
    details = dict(field.split('=', 1) for field in detail.split(';')[1:])
    assert details['calls'] == '16'
    count = b['value0']
    expected = count if details['order'] == 'descending' else 1
    assert count in (4, 64, 256) and c['value0'] == count and b['value1'] == c['value1'] == expected
    assert b['unit'] == c['unit'] == 'us' and cap['unit'] == 'count'
    assert all(row['hr'] == 0 for row in (b, c, cap)) and cap['value1'] == 0 and cap['value0'] > 0
    pair = int(details['pair'])
    first = 'baseline' if ((pair % 2 == 0) != (details['storage'] == 'descending')) else 'candidate'
    assert details['first'] == first
    pairs.append({'tasks': count, 'storage': details['storage'], 'order': details['order'], 'pair': pair, 'first': first,
                  'calls_per_batch': 16, 'expected_id': expected, 'baseline_batch_us': b['durationUs'], 'candidate_batch_us': c['durationUs'],
                  'baseline_collection_capacity_bytes': cap['value0'], 'candidate_collection_capacity_bytes': cap['value1']})
assert len({(p['tasks'], p['storage'], p['order'], p['pair']) for p in pairs}) == 48
with (PACKET/'Debug/queue-pairs.csv').open('w', encoding='utf-8', newline='') as handle:
    writer = csv.DictWriter(handle, fieldnames=list(pairs[0]))
    writer.writeheader(); writer.writerows(pairs)
summaries = []
for key in sorted({(p['tasks'], p['storage'], p['order']) for p in pairs}):
    cohort = [p for p in pairs if (p['tasks'], p['storage'], p['order']) == key]
    b, c = [p['baseline_batch_us'] for p in cohort], [p['candidate_batch_us'] for p in cohort]
    summaries.append({'tasks': key[0], 'storage': key[1], 'order': key[2], 'pairs': len(cohort),
                      'baseline_batch_us': {'samples': b, 'median': statistics.median(b), 'min': min(b), 'max': max(b)},
                      'candidate_batch_us': {'samples': c, 'median': statistics.median(c), 'min': min(c), 'max': max(c)},
                      'candidate_minus_baseline_median_us': statistics.median(c) - statistics.median(b),
                      'baseline_collection_capacity_bytes': sorted({p['baseline_collection_capacity_bytes'] for p in cohort}),
                      'candidate_collection_capacity_bytes': [0]})
write_json(PACKET/'Debug/component-summary.json', {
    'status': 'DEBUG DIAGNOSTIC COMPONENT OBSERVATIONS ONLY; Release pending; whole repository not evaluated',
    'queue_protocol': {'pairs': 48, 'duration_rows': 96, 'capacity_rows': 48, 'calls_per_batch': 16, 'cohorts': summaries},
    'guard_controls': [{'detail': r['detail'], 'diagnostic_us': r['durationUs'], 'prompt_count': r['value0'], 'discard_offered': r['value1'], 'hr': r['hr']} for r in guards],
    'power_controls': [{'control': r['value0'], 'name': power_names[r['value0']], 'wait_calls': r['value1'], 'hr': r['hr']} for r in power],
    'interpretation': 'No p95, historical application speedup, allocator-call, heap-peak, queue-start or Release non-regression claim.',
})

queue_path = 'RedSalamander/FolderWindow.FileOperations.State.Queue.cpp'
legacy = git('show', f'{BASE}:{queue_path}').decode('utf-8').replace('\r\n', '\n')
match = re.search(r'    const auto firstUnblockedQueuedTaskId = \[&\]\(\) noexcept -> uint64_t\n    \{\n.*?\n    \};\n', legacy, re.S)
assert match and sha(match.group(0).encode()) == BASE_HASH
source_dir = PACKET/'source-excerpts'
source_dir.mkdir()
(source_dir/'baseline-selector-78a8a75c4.txt').write_text(match.group(0), encoding='utf-8', newline='\n')
current_queue = git('show', f'{HEAD}:{queue_path}').decode('utf-8').replace('\r\n', '\n')
def extract_between(text, start, end): return text[text.index(start):text.index(end, text.index(start))]
excerpts = {
    'candidate-selector-e83.txt': extract_between(current_queue, 'uint64_t FolderWindow::FileOperationState::FirstUnblockedQueuedTaskIdLocked()', '\n#ifdef ENABLE_TESTS'),
    'baseline-adapter-e83.txt': extract_between(current_queue, 'uint64_t FolderWindow::FileOperationState::DebugBaselineQueueHeadLocked(', '\n#endif'),
    'queue-witness-e83.txt': extract_between(current_queue, '    // queue_head_selector_v1:', '\nbool FolderWindow::FileOperationState::DebugValidateInterlockWaitingPresentationForSelfTest'),
}
for name, content in excerpts.items(): (source_dir/name).write_text(content, encoding='utf-8', newline='\n')
source_paths = [queue_path, 'RedSalamander/FolderWindow.FileOperations.State.cpp', 'RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.ItemResultReducer.cpp', 'Common/Helpers.h', 'Specs/Testing/Testing_PerformanceValidation.md', 'Specs/FileSystem/FileSystem_FileOperations.md']
write_json(PACKET/'source-excerpts/provenance.json', {
    'candidate_commit': HEAD, 'baseline_commit': BASE,
    'baseline_path': queue_path, 'baseline_git_blob': git('rev-parse', f'{BASE}:{queue_path}').decode().strip(),
    'baseline_canonical_lf_utf8_excerpt_sha256': BASE_HASH,
    'candidate_git_blobs': {name: git('rev-parse', f'{HEAD}:{name}').decode().strip() for name in source_paths},
    'adaptations': 'Frozen old lambda wrapped as a private member; one real capacity observation after collection. Current harness uses symmetric ID plus relaxed atomic capacity outputs, defaulted/deleted sample lifecycle, actual production selector helper.',
    'observer_caveats': 'Common scalar/atomic output stores, mutex, timer, compiler code-generation and baseline capacity/local adapter costs remain included. Tiny deltas cannot be attributed solely to selector logic. Fixtures/assertions/format/JSONL output are outside spans.',
    'excerpt_hashes': {p.name: sha(p.read_bytes()) for p in source_dir.glob('*.txt')},
})

write_json(PACKET/'evidence-bindings.json', {
    'schema': 'red-salamander.native-component-evidence-preparation.v1',
    'status': 'DRAFT DEBUG SUBSECTION; NOT ARCHIVE-VALIDATED; RELEASE PENDING; NO FULL/PRODUCTION QUALIFICATION',
    'prepared_utc': datetime.now(timezone.utc).isoformat(), 'receipt_id': RECEIPT_ID, 'receipt_file_sha256': sha(receipt_bytes),
    'candidate_source': HEAD, 'candidate_snapshot': SNAPSHOT, 'machine_profile': '4cb089111a23',
    'proposed_archive_path': 'Specs/TestRuns/4cb089111a23/FileOps/2026-11-13_165916_NativeComponents',
    'historical_receipt_selected_outputs': selected_output_checks, 'runs': bindings, 'exact_copies': copies,
    'raw_preservation': 'Both complete metric streams remain unchanged in their sealed run roots. Packet JSONL is an exact ordered metric-name excerpt, with one-based original line indexes and per-line SHA256; no commit metadata is filled in.',
    'release': {'status': 'PENDING', 'receipt': None, 'runs': [], 'measurements': None},
})

write_json(PACKET/'Release/PENDING.json', {'status': 'PENDING; no Release build/run/receipt/measurement in this packet',
    'required': 'After current Affected native operation reaches its terminal owned-process quiet point, use fresh exact test-enabled Release receipt and governed exact noactivation filters. Preserve every attempt, same source/protocol/fixture/machine, new process IDs and hashes; append rather than relabel Debug rows.'})

readme = '''# Native component evidence — draft Debug subsection

Status: **Debug diagnostic evidence prepared; Release and archive validation pending.**
Source: `e83dd755f44cef4f720d3f439f5fbd06383a569a`; snapshot and exact build receipt are bound in
`evidence-bindings.json`. No native metric's blank commit field was filled in. This packet joins the
outer exact filtered invocation, matching after/final workspace identities, receipt, executable path,
owned native process/exit record, metric timestamps and byte hashes. Sealed prelaunch files rotated
into `previous_run` explicitly bind the receipt ID, source snapshot, epoch zero and exact filtered
validation plan; archive inventory binds those bytes. Their original fingerprint/plan digests remain
intact rather than being reinterpreted as v2 promotions. The immutable receipt exactly matches its
copy in the sealed full-Debug-build context. Its executable/FileSystem.dll/Common.dll output hashes
are historical receipt identities, not fresh live-file checks: the authorized active Affected writer
is rebuilding/re-attesting this same Debug profile, so mutable current bytes are not historical proof.

Both selected FileOps runs have three native passes (Setup, requested case, Cleanup), zero failures,
zero skips, native process exit zero and suite PASSED. Outer wrappers exit 2 / repository
NOT_EVALUATED. Preserve their non-clean **non-blocking** sandbox audits: Review 100 issues; Identity
99 issues (19 root-child, 8 stale-run, 72 manual/unparseable sibling-run). Exact original lists/reasons
remain in the copied aggregate and `sandbox-audit.json`; ownership of unrelated residue is unproven.
These are not current-run root-escape/reparse findings, a clean-root claim or permission to clean.
Owning contract: `Testing_SelfTests.md` aggregate-audit/non-follow/preflight sections.

## Queue selector component

`Debug/Identity/perf/perf_metrics.jsonl` retains all 48 paired groups: 96 duration rows plus 48 capacity
rows. All three fixture sizes (4/64/256), storage/key directions and four counterbalanced pair indexes
are independently checked by `prepare_packet.py`; `Debug/queue-pairs.csv` retains every observation,
and `component-summary.json` includes all samples, medians, ranges and positive/negative deltas.
Durations measure **16-call batches**, not whole queue latency. Capacity's raw JSONL unit is `count`;
the protocol defines `value0` as baseline collection capacity bytes and `value1` as candidate bytes.
This resource observation is neither allocator calls nor heap/process peak. Correct selected IDs and
every observed scalar were asserted by the native case; packet checks additionally match expected
fixture ID, pair order, successful HRESULTs, actual raw rows and zero candidate collection capacity.

The actual production helper is compared with the frozen old algorithm in the same executable.
`source-excerpts/` binds the immutable baseline excerpt and current helper/adapter/harness. Both paths
retain matched per-call ID plus relaxed atomic capacity outputs. Baseline capacity-read/local adapter,
common output stores, nested task mutex, timing and compiler costs remain included. Tiny differences
cannot be assigned solely to selector logic. No p95 threshold or historical application speedup is claimed.
Debug is diagnostic; no Release latency, NAS, queue-start, allocation-saving or whole-queue acceptance.

## Guard and power witnesses

Review includes three real private admission/cleanup controls, each with one prompt and zero Delete
source offers: namespace-different IDs for the same physical entry (Cancel), distinct bound files
(Cancel), and distinct identityless overwrite (publication plus typed Completed/Retained). Exact bytes
and independent positive identity refusal were asserted natively. The namespace seam models a reported
ID-domain mismatch while retaining physical handles; it is not physical NAS/export evidence. The policy
conservatively withholds destructive distinctness for differing qualified Local roots, even real distinct
files or GUID-looking roots. Ordinary bound cleanup has separate source/publication authority.

All six actual conflict-wait controls are retained: Preanswered, Precancelled, QuietCancel, QuietResume,
StopOnly and Spurious. Counts describe bounded event/deadline calls under the test-only short quiet
budget; fake power APIs assert exact errors and failed-clear retirement. Private unadmitted tasks neither
exercise real Windows power/sleep nor measure historical polling CPU. Timings include declared responder
drain, and added correctness scenario durations are not equal-work historical comparisons.

## Preservation and remaining steps

Raw metrics exceed the 2-MiB file / 5-MiB run Git archive caps and remain in sealed local evidence.
Each compact JSONL preserves exact original row bytes/order; adjacent indexes bind full raw SHA256,
byte/line counts and every retained original line. Archive intent/completion records bind those raw
streams. Exact result/trace/receipt/workspace copies preserve original bytes and terminal newlines.
`Release/PENDING.json` is an honest empty slot, not a complete report.

Before publication: finish current native operation, obtain and review matched-source test-enabled
Release evidence, retain failures/skips and all samples, review this packet, add an exact archive
`-text -whitespace` attribute rule (currently text:auto would normalize digest-bound rows), then validate
the **complete** candidate packet with `Tools/Test-TestRunArchive.ps1 -RunPath <packet>` before any
force-add. Final closeout also requires the owning inventory validation. Neither validator has been
run by this preparation task. Do not claim Full/ARM64/ASan/physical-device/interactive qualification.
No source files, Specs archive, builds, tests or current active native run were changed by this task.
'''
(PACKET/'README.md').write_text(readme, encoding='utf-8', newline='\n')

# Publish a deterministic recipe and a whole-packet membership/digest list without self-recursive hashes.
recipe = '''# Preparation recipe (read-only source / scratch-only writes)
1. Create a fresh explicit New-RSTestRunContext with RunId='' and existing marked Z root; do not initialize roots.
2. Run prepare_packet.py in that owned scratch. It verifies expected raw hashes, sealed archive transactions,
   result counts, prelaunch fingerprint/workspace/receipt/PID/time bindings, exact metric rows and protocol groups.
3. Inspect Debug/component-summary.json, every regression/range and the complete original audit objects.
4. Preserve whole raw metric files in their sealed locations; never replace them with this excerpt.
5. Future Release and complete archive checks are separate parent-owned steps; nothing here runs native tests.
'''
(PACKET/'recipe.txt').write_text(recipe, encoding='utf-8', newline='\n')
files = [{'path': str(p.relative_to(PACKET)).replace('\\','/'), 'bytes': p.stat().st_size, 'sha256': sha(p.read_bytes())}
         for p in sorted(PACKET.rglob('*')) if p.is_file()]
assert all(item['bytes'] <= 2*1024*1024 for item in files)
assert sum(item['bytes'] for item in files) <= 5*1024*1024
write_json(SCRATCH/'packet-manifest.json', {'packet': str(PACKET), 'status': 'PREPARED, NOT ARCHIVE-VALIDATED', 'files': files,
    'file_count': len(files), 'total_bytes': sum(item['bytes'] for item in files), 'source_unchanged_after': not git('status','--porcelain').strip()})
assert not git('status', '--porcelain').strip()
print(json.dumps({'packet': str(PACKET), 'files': len(files), 'total_bytes': sum(item['bytes'] for item in files),
                  'raw_and_selected_counts': {name: {'raw': b['raw_metrics_rows'], 'excerpt': b['excerpt_metrics_rows'], 'audit': b['sandbox_audit']['category_counts']} for name,b in bindings.items()},
                  'queue_pairs': len(pairs), 'guard_rows': len(guards), 'power_rows': len(power),
                  'queue_debug_median_deltas_us': [row['candidate_minus_baseline_median_us'] for row in summaries]}, indent=2))
