import csv
import hashlib
import json
from pathlib import Path

repo = Path('C:/Users/PVB/.codex/worktrees/viewer-space-explorer-analyst/RedSalamander')
evidence = Path('D:/RedSalamander.Perf/evidence/viewer-space-finish-all-20261009')
assert (evidence.parents[1] / '.red-salamander-test-root').read_text().strip() == 'red-salamander.test-sandbox-root.v1'
archive = repo / 'Specs/TestRuns/7d3a1247382a/Validation/space_release_prepush_commands_captures_20261009T213050Z'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
rows = list(csv.DictReader((archive / 'manifest.tsv').open(encoding='utf-8-sig'), delimiter='\t'))
replay = []
for row in rows:
    p = archive / row['path']
    assert p.stat().st_size == int(row['bytes']) and sha(p) == row['sha256']
    raw = Path(row['absolute_source'])
    if raw.exists():
        assert raw.stat().st_size == int(row['bytes']) and sha(raw) == row['sha256']
    replay.append({'path': row['path'], 'sha256': row['sha256'], 'bytes': int(row['bytes']), 'raw_available': raw.exists()})
assert len(rows) == 12
provenance = archive / 'selftest/last_run/perf/viewer_space_captures/provenance.json'
data = json.loads(provenance.read_text(encoding='utf-8-sig'))
assert data['scenario'] == 'cmd_viewer_space_explorer_analyst_exact_nested_detail'
assert data['captureKind'] == 'actual-application-harness' and data['noActivation'] is True
assert data['dpi'] == 144 and len(data['frames']) == 11
assert {f['file'] for f in data['frames']} == {Path(r['path']).name for r in rows if r['path'].endswith('.png')}
assert all(f['layoutGeneration'] == f['presentedLayoutGeneration'] and f['layoutDirty'] is False for f in data['frames'])
receipt = json.loads((repo / '.build/x64/Release/build-receipt.json').read_text(encoding='utf-8-sig'))
assert receipt['receipt_id'] == '37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd'
assert receipt['source_snapshot_id'] == '9addaf53d4d88d618223a7f2d8554ac0b73f2472946c3a3e0875b1fd94ac9554' and receipt['tests_enabled'] is True
binaries = []
for relative in ['.build/x64/Release/RedSalamander.exe', '.build/x64/Release/Plugins/ViewerSpace.dll']:
    output = next(o for o in receipt['outputs'] if o['relative_path'] == relative)
    p = repo / relative
    assert p.stat().st_size == output['size_bytes'] and sha(p) == output['sha256']
    binaries.append({'path': relative, 'sha256': output['sha256'], 'bytes': output['size_bytes']})
review = {
    'reviewer': 'root', 'scope': 'whole-frame visual inspection of eleven original-resolution harness frames',
    'run_id': '20261009T213050Z-43264-4bde93ea855e44c885a3277a68e89038',
    'receipt_id': receipt['receipt_id'], 'build_source_snapshot_id': receipt['source_snapshot_id'],
    'capture_provenance_sha256': sha(provenance), 'payload_replay': replay, 'binaries': binaries,
    'observations': [
        'Explorer light displays nested folder header bands and file folded-corner markers, a ranked grid, and selected file details.',
        'Analyst dark, Rainbow, and high-contrast display the nested map above the five-column grid with an adjacent selection inspector.',
        'Rainbow preserves file/folder shape differences and selection outline.',
        'Narrow Explorer stacks the grid and inspector beneath the map; grid horizontal scrolling preserves additional columns.',
        'At actual width 1100 DIP analysis is right; at 1099.33 DIP it is bottom.',
        'At actual width 504 DIP analysis is visible; at 503.33 DIP it is hidden and the map fills the body.',
        'At available body height 456 DIP analysis is visible; at 455.33 DIP it is hidden and the map fills the body.',
        'Small rectangles elide or omit text while retaining geometry and file corner markers; selected file name and byte count remain available in the inspector when analysis is visible.'
    ],
    'limits': 'These are deterministic English 144-DPI frames. This record does not qualify other DPI, Narrator, human contrast usability, live providers, or performance targets.'
}
out = evidence / 'root-release-commands-capture-review.json'
out.write_text(json.dumps(review, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'record': str(out), 'payloads_verified': len(replay), 'frames_reviewed': len(data['frames']), 'raw_available': sum(r['raw_available'] for r in replay), 'binaries_verified': len(binaries)}))
