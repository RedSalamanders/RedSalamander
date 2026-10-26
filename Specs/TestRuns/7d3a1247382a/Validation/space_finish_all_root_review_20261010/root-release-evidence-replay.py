import csv
import hashlib
import json
from pathlib import Path

repo = Path('C:/Users/PVB/.codex/worktrees/viewer-space-explorer-analyst/RedSalamander')
evidence = Path('D:/RedSalamander.Perf/evidence/viewer-space-finish-all-20261009')
assert (evidence.parents[1] / '.red-salamander-test-root').read_text().strip() == 'red-salamander.test-sandbox-root.v1'
base = repo / 'Specs/TestRuns/7d3a1247382a'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
archives = ['Build/space_candidate_release_prepush_reattest_20261009T214710Z', 'Validation/space_release_prepush_reports_20261009T213050Z', 'Validation/space_release_prepush_native_captures_20261009T213050Z']
results = []
for name in archives:
    archive = base / name
    rows = list(csv.DictReader((archive / 'manifest.tsv').open(encoding='utf-8-sig'), delimiter='\t'))
    for row in rows:
        p = archive / row['path']
        assert p.stat().st_size == int(row['bytes']) and sha(p) == row['sha256']
        raw = Path(row['absolute_source'])
        if raw.exists():
            assert raw.stat().st_size == int(row['bytes']) and sha(raw) == row['sha256']
    results.append({'archive': name, 'pairs_verified': len(rows), 'raw_available': sum(Path(r['absolute_source']).exists() for r in rows)})
receipt = json.loads((repo / '.build/x64/Release/build-receipt.json').read_text(encoding='utf-8-sig'))
assert receipt['receipt_id'] == '37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd'
exe = next(o['sha256'] for o in receipt['outputs'] if o['relative_path'] == '.build/x64/Release/ViewerPETests.exe')
dll = next(o['sha256'] for o in receipt['outputs'] if o['relative_path'] == '.build/x64/Release/Plugins/ViewerSpace.dll')
native = base / archives[-1]
frames = []
for p in sorted(native.glob('viewer-pe/viewer-space-lifecycle-*/*.provenance.json')):
    d = json.loads(p.read_text(encoding='utf-8-sig'))
    assert d['valid'] is True and d['foreground_preserved'] is True and d['capture_hresult'] == 0
    assert d['capture_scope'] == 'owned WS_EX_NOACTIVATE parent with embedded ViewerSpace child'
    assert d['dpi'] == 144 and d['executable_sha256'] == exe and d['viewer_space_dll_sha256'] == dll
    assert d['last_render_present_hr_after'] == 0
    png = p.with_name(p.name.replace('.provenance.json', '.png'))
    assert png.is_file()
    frames.append({'scenario': d['scenario'], 'png': str(png.relative_to(native)), 'png_sha256': sha(png), 'provenance_sha256': sha(p)})
assert len(frames) == 8
focused = base / 'Viewers/space_debug_exact_activation_passed_20261009T223523Z'
focus_pairs = []
for line in (focused / 'manifest.txt').read_text().splitlines():
    if not line or line.startswith('#'):
        continue
    expected, size, name, source = line.split('  ', 3)
    p = focused / name
    assert p.stat().st_size == int(size) and sha(p) == expected
    raw = Path(source)
    assert raw.stat().st_size == int(size) and sha(raw) == expected
    focus_pairs.append({'path': name, 'sha256': expected, 'bytes': int(size)})
assert len(focus_pairs) == 2
fd = json.loads((focused / 'viewer-space-focused-diagnostic.json').read_text(encoding='utf-8-sig'))
assert fd['post_run_attestation_intact'] is True and fd['repository_verdict'] == 'NOT_EVALUATED'
assert fd['build_receipt_id'] == 'f721d4c3644b36016e3b81900978cc335f1a530d9eadc094f585683bc7f3638b'
assert fd['launcher_sha256'] == sha(repo / '.build/run-viewer-space-native.ps1')
assert fd['canonical_runner_sha256'] == sha(repo / 'Tools/Run-AllTests.ps1')
assert len(fd['results']) == 1 and fd['results'][0]['exit_code'] == 0 and fd['results'][0]['desktop_lease'] is True
case = 'TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses'
log = (focused / (case + '.output.log')).read_text(encoding='utf-8-sig')
for witness in ['fresh harness process reports an exit code', 'fresh harness run passes', 'owned job has no active processes before desktop restoration', 'parent observes exact saved foreground and keyboard focus without restoration writes']:
    assert '[PASS] ' + case + ' ' + witness in log
assert '[FAIL]' not in log
record = {
    'reviewer': 'root', 'archives': results, 'lifecycle_frames_reviewed_at_original_resolution': frames,
    'lifecycle_observations': 'Before data: spinner and processing path. Partial scan: observed totals, completed files and scanning folder row. Done: nested file fills folder. Refresh pending: old map and rows with spinner and refresh state. Cancel: old completed totals/map restored. Refused: explicit memory status keeps old map. Empty: Scan completed. Provider error: Error status.',
    'extra_inflight_frame': 'Viewed the additional 96-DPI text-sidecar frame. It shows scanning/processing/spinner but does not supply the complete lifecycle JSON provenance or other-DPI qualification.',
    'focused_debug_pairs': focus_pairs, 'focused_debug_scope': 'exact ViewerSpace activation case, repository NOT_EVALUATED',
    'limits': 'Whole-frame review is limited to deterministic English 144-DPI lifecycle captures. Broad Release repository verdict remains FAILED. No target or other-platform/accessibility qualification is implied.'
}
out = evidence / 'root-release-native-reports-build-and-debug-focus-replay.json'
out.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'record': str(out), 'pairs_verified': sum(r['pairs_verified'] for r in results) + len(focus_pairs), 'lifecycle_frames_reviewed': len(frames), 'focused_debug': 'PASSED'}))
