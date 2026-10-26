import hashlib
import json
import subprocess
from pathlib import Path

root = Path(__file__).parent
repo = Path('Z:/src/RedSalamander')
packet = repo / 'Specs/TestRuns/4cb089111a23/Validation/2026-09-27_default-family-dispatch'
assert not packet.exists(), 'Retain completed evidence; never overwrite it'
definitions = [
    ('v5-release', 'fileops-current-v5-full-20260927', 'fileops', 1),
    ('v6-release', 'family-dispatch-v6-release-20260927', 'fileops', 0),
    ('v6-debug', 'family-dispatch-v6-debug-20260927', 'fileops', 0),
    ('v6-debug-common', 'family-dispatch-v6-debug-20260927', 'common_smoke', 0),
]

def read_json(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))

def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

# Verify every complete immutable runner copy before producing a reviewed subset.
old_inputs = {row['path']: row['sha256'] for row in read_json(root / definitions[0][1] / 'candidate-inputs.json')}
new_inputs = {row['path']: row['sha256'] for row in read_json(root / definitions[1][1] / 'candidate-inputs.json')}
debug_inputs = {row['path']: row['sha256'] for row in read_json(root / definitions[2][1] / 'candidate-inputs.json')}
assert len(old_inputs) == 36 and old_inputs.keys() == new_inputs.keys()
assert new_inputs == debug_inputs
assert [name for name in old_inputs if old_inputs[name] != new_inputs[name]] == ['RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp']
for dirname in (definitions[1][1], definitions[2][1]):
    assert int((root / dirname / 'build-exit.txt').read_text(encoding='utf-8-sig')) == 0
    assert 'Diagnostics: 0 warning(s), 0 error(s)' in (root / dirname / 'build.log').read_text(encoding='utf-8-sig')
analysis = []
for label, dirname, key, expected_exit in definitions:
    source = root / dirname
    build_source = root / 'exact-bound-results-candidate-v5-20260927' if label == 'v5-release' else source
    build = read_json(build_source / 'build-receipt.json')
    assert build['git_head'] == 'ccebd5af68eb45f2e9deed64a7063780e71eb1ce'
    assert build['platform'] == 'x64' and build['configuration'] == ('Debug' if 'debug' in label else 'Release')
    assert build['exit_code'] == 0 and build['tests_enabled'] is True
    assert build['target'] == 'solution' and build['project_selection'] == ['RedSalamander.sln']
    assert f"Skipping build (verified receipt {build['receipt_id']})" in (source / f'{key}.log').read_text(encoding='utf-8-sig')
    assert int((source / f'{key}-exit.txt').read_text(encoding='utf-8-sig')) == expected_exit
    manifests = read_json(source / f'{key}-runner-copy-hashes.json')
    archive = source / f'{key}-runner-artifacts'
    for entry in manifests:
        assert sha(archive / entry['path']) == entry['sha256'], entry['path']
    result = read_json(archive / 'artifacts/selftest/last_run/fileops/results.json')
    skipped = [case for case in result['cases'] if case['status'] == 'skipped']
    unexpected = [case for case in skipped if case.get('skip_class') == 'unexpected']
    assert result['failed'] == 0
    if label == 'v5-release':
        assert len(unexpected) == 1 and unexpected[0]['name'] == 'FileOps_BatchRenameExactResultCarrier'
        assert unexpected[0]['reason'] == 'not executed'
    else:
        assert not unexpected
        if key == 'common_smoke':
            common = next(case for case in result['cases'] if case['name'] == 'FileOps_CommonPresentationPerf')
            assert common['status'] == 'passed'
        if key == 'fileops':
            batch = next(case for case in result['cases'] if case['name'] == 'FileOps_BatchRenameExactResultCarrier')
            assert batch['status'] == 'passed'
            recycle = next(case for case in result['cases'] if case['name'] == 'Phase10_DeferredConsentAndRecycleEscalation')
            assert recycle['status'] == 'passed'
    assert 'Disk audit issues: 0' in (source / f'{key}.log').read_text(encoding='utf-8-sig')
    analysis.append(dict(label=label, runner_exit=expected_exit, files_rehashed=len(manifests),
                         build_receipt_id=build['receipt_id'], configuration=build['configuration'],
                         passed=result['passed'], failed=result['failed'], skipped=result['skipped'],
                         skips=skipped, native_started_utc=result['run_started_utc'],
                         artifact_root=result['artifact_root']))

packet.mkdir(parents=True)
mapping = []

def copy(source, relative):
    target = packet / relative
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(source.read_bytes())
    mapping.append(dict(source=str(source), path=relative, sha256=sha(source), bytes=source.stat().st_size))

for label, dirname, key, expected_exit in definitions:
    source = root / dirname
    prefix = f'raw/{label}/'
    for name in (f'{key}-exit.txt', f'{key}-plan.json', f'{key}-runner-copy-hashes.json', f'{key}.log', 'invocations.json', 'candidate-inputs.json'):
        copy(source / name, prefix + name + ('.receipt.txt' if name.endswith('.json') else ''))
    last = source / f'{key}-runner-artifacts/artifacts/selftest/last_run'
    for name in ('results.json', 'run-all-tests-results.json', 'workspace-final.json', 'workspace-after-selftest.file-operations.json', 'trace.txt', 'fileops/results.json', 'fileops/trace.txt'):
        copy(last / name, prefix + name + ('.receipt.txt' if name.endswith('.json') else ''))
    for name in ('run.ps1', 'run-focused.ps1', 'build-and-test.ps1', 'build.log', 'build-exit.txt', 'build-receipt.json', 'run-plan-decision.json'):
        if (source / name).exists() and label != 'v6-debug-common':
            copy(source / name, prefix + name + ('.receipt.txt' if name.endswith('.json') else ''))

copy(root / 'exact-bound-results-candidate-v5-20260927/build-receipt.json', 'raw/v5-release/build-receipt.json.receipt.txt')
copy(root / 'family-dispatch-v6-sequence-20260927.ps1', 'raw/sequence.ps1')
copy(Path(__file__), 'raw/curate-family-dispatch.py')
source_name = 'RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp'
copy(root / 'family-dispatch-v6-release-20260927/inputs' / source_name, 'raw/family-registration.cpp.source.txt')
patch = subprocess.check_output(['git', 'diff', '--', source_name], cwd=repo)
assert b'Every default FileOps phase must be dispatched by a child family.' in patch
(packet / 'raw/family-registration.patch').write_bytes(patch)
(packet / 'raw/analysis.receipt.txt').write_text(json.dumps(analysis, indent=2) + '\n', encoding='utf-8')
(packet / 'raw/file-map.receipt.txt').write_text(json.dumps(mapping, indent=2) + '\n', encoding='utf-8')
(packet / '.gitattributes').write_bytes(b'raw/** -text -whitespace\nSHA256SUMS.txt -text -whitespace\n')
rows = '\n'.join(f"| {r['label']} | {r['passed']} / {r['failed']} / {r['skipped']} | {r['runner_exit']} | {r['files_rehashed']} |" for r in analysis)
(packet / 'README.md').write_bytes(f'''# Default File Operations family dispatch

The broad Release runner exposed a test-registration defect: the new exact Batch Rename phase passed when selected directly but was missing from the default child-family table. V5's application process completed normally and reported the case as an unexpected `not executed` skip. The governed wrapper correctly rejected it with `FORBIDDEN_SKIP_CLASS, UNALLOWED_SKIP` and exit 1. This failed run remains unchanged.

V6 registers the dedicated Batch Rename family and adds a compile-time invariant requiring every default phase except setup/cleanup to belong to a dispatched family. Existing family overlap remains valid. Only 26 lines in the test registration translation unit change; production and all other 35 frozen source inputs remain unchanged. The source, patch, frozen input manifests and full-solution build receipts are retained. Both native profiles build with zero warnings/errors.

| Run | Native passed / failed / skipped | Wrapper exit | Copied files independently rehashed |
| --- | ---: | ---: | ---: |
{rows}

Both V6 broad runs execute the Batch Rename case successfully through default family dispatch. Each also executes the real Recycle Skip carrier case. Debug additionally runs the opt-in common presentation fixture; this is a runtime smoke, not a paired timing result. All four disk audits have zero findings. Release's declared skips cover its Debug-only scheduler export, two 7z scenarios owned by Compare coverage, and absent remote profiles; Debug's exact skips are retained in the analysis and native receipts. Declared skips are not treated as executed capability passes.

The Debug run was strengthened from three focused cases to broad FileOps plus the common smoke before its build started. Its timestamped decision and driver are retained. A single sequential lane owns all builds/tests, uses the governed nonactivating runner, and explicitly authorizes the marked alternate-volume test root. No foreground interaction or screenshot is performed by this change.

The isolated consumer remains on base `ccebd5af68eb45f2e9deed64a7063780e71eb1ce`, with the exact 36 overlays and qualified DxUi pin `78b3de389a189c7f86f611787e0489fb6d474218`. The candidate grid/menu branches are not adopted. These results qualify the current x64 Debug/Release FileOps suite slice only. Full runner, foreground journeys, UIA/AT, final combined pin, native monitor transitions/192 DPI, final gallery and integrated resources remain separate gates. No performance exception is accepted here.

Full immutable runner trees, including complete large telemetry streams, remain beneath `C:/RedSalamander.Perf/evidence/i26-ui/` in the three directories named by the retained curator. Their manifests bind all files; the reviewed packet copies complete claim-bearing results, traces, workspace/source snapshots and runner logs without truncation. Raw JSON uses `.receipt.txt` suffixes without changing bytes. `SHA256SUMS.txt` covers every payload; `raw/file-map.receipt.txt` maps copies to their sources.
'''.encode('utf-8'))
files = sorted(path for path in packet.rglob('*') if path.is_file())
assert max(path.stat().st_size for path in files) < 2 * 1024 * 1024
assert sum(path.stat().st_size for path in files) < 5 * 1024 * 1024
(packet / 'SHA256SUMS.txt').write_bytes(''.join(sha(path) + '  ' + path.relative_to(packet).as_posix() + '\n' for path in files).encode('ascii'))
print(json.dumps(dict(payloads=len(files), bytes=sum(path.stat().st_size for path in files), results=analysis), indent=2))
