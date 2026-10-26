import hashlib
import json
import pathlib
import struct

repo = pathlib.Path(__file__).resolve().parents[2]
authority = pathlib.Path('D:/RedSalamander.Perf')
assert (authority / '.red-salamander-test-root').read_text().strip() == 'red-salamander.test-sandbox-root.v1'
evidence = authority / 'evidence/viewer-space-finish-all-20261009'
sidecar = repo / 'docs/res/viewer-space-captures.provenance.json'
document = json.loads(sidecar.read_bytes())
captures = [*document['releaseExplorerAnalystCapture']['captures'], {**document['nativeReleaseInflightCapture'], 'file': 'viewer-space-inflight.png'}]
assert len(captures) == 3 and document['historicalCaptures']
pins = {
    'viewer-space-explorer.png': (503887, 'd34452778aa6413f65b8df79446db6c37fa7a2e707aca9c6c3b0618f87b416d7'),
    'viewer-space-analyst-rainbow.png': (383793, '3960a081a380ee9708c3e9cc7564542e9a811bc0ffaf96066abda089c3bdf41b'),
    'viewer-space-inflight.png': (107897, '828e81d84b269d22079bb8a179b130dabac72d1f8fe284d30aa587b5466c66c1'),
}
checked = []
for capture in captures:
    name = capture['file']
    data = (repo / 'docs/res' / name).read_bytes()
    assert len(data) == pins[name][0] == capture['bytes']
    digest = hashlib.sha256(data).hexdigest()
    assert digest == pins[name][1] == capture['sha256']
    source = (repo / capture['source']).resolve(strict=True)
    assert source.is_relative_to((repo / 'Specs/TestRuns').resolve())
    assert data == source.read_bytes()
    assert data[:8] == b'\x89PNG\r\n\x1a\n'
    width, height = struct.unpack('>II', data[16:24])
    assert width == capture['pngPixelDimensions']['width'] and height == capture['pngPixelDimensions']['height']
    checked.append({'file': name, 'source': capture['source'], 'bytes': len(data), 'sha256': digest, 'pngWidth': width, 'pngHeight': height})
for key in ['releaseExplorerAnalystCapture', 'nativeReleaseInflightCapture']:
    build = document[key]['build']
    assert build['receiptId'] == '37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd'
    assert build['buildSnapshotId'] == '9addaf53d4d88d618223a7f2d8554ac0b73f2472946c3a3e0875b1fd94ac9554'
    assert build['testsEnabled'] and build['configuration'] == 'Release' and build['platform'] == 'x64'
old = repo / 'Specs/TestRuns/7d3a1247382a/Validation/space_finish_all_root_review_20261010/root-doc-provenance-before-final.json'
assert hashlib.sha256(old.read_bytes()).hexdigest() == 'f43758ab71c3695860165c945fea73eafe73cf7aa0095f1b598c3b8f32ef085b'
record = {'reviewer': 'root', 'mismatches': 0, 'captures': checked,
          'sidecarSha256': hashlib.sha256(sidecar.read_bytes()).hexdigest(),
          'historicalProvenancePreserved': True, 'scope': 'Three byte-exact accepted harness captures from named frozen-source144DPI Release scenarios; no screenshot generation or new DPI/platform/accessibility qualification.'}
with (evidence / 'root-current-documentation-capture-replay.json').open('x', encoding='utf-8') as stream:
    json.dump(record, stream, indent=2)
    stream.write('\n')
print(json.dumps(record))
