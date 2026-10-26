import hashlib
import importlib.util
import json
import pathlib
import sys

spec = importlib.util.spec_from_file_location('space_preservation_guard', pathlib.Path(__file__).with_name('archive-current-space-campaign.py'))
guard = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = guard
spec.loader.exec_module(guard)
guard._check_marker_first()
authority = pathlib.Path('D:/RedSalamander.Perf').resolve()
destination = authority / 'evidence/viewer-space-finish-all-20261009/hover-navigation-current'
guard._safe_source_path(destination / 'request.json', authority)
request = json.loads((destination / 'request.json').read_text(encoding='utf-8-sig'))
assert request['finishedUtc'] and request['sourceSnapshotId'] == '9addaf53d4d88d618223a7f2d8554ac0b73f2472946c3a3e0875b1fd94ac9554'
assert request['receiptId'] == '37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd'
run_id = request['runId']
assert '/' not in run_id and '\\' not in run_id and '..' not in run_id
source = authority / 'runs' / run_id / 'artifacts'
guard._check_existing_ancestors(source)
assert source.is_dir() and source.resolve().is_relative_to(authority)
target = destination / 'artifacts'
assert not target.exists()
guard._check_existing_ancestors(target, allow_missing_leaf=True)
files = []
for path in sorted(source.rglob('*')):
    guard._safe_source_path(path, authority)
    if path.is_file():
        files.append(path)
assert 1 <= len(files) <= 2000
target.mkdir()
pins = []
for path in files:
    relative = path.relative_to(source)
    output = target / relative
    assert output.resolve().is_relative_to(target.resolve())
    output.parent.mkdir(parents=True, exist_ok=True)
    data = path.read_bytes()
    with output.open('xb') as stream:
        stream.write(data)
    assert output.read_bytes() == data == path.read_bytes()
    pins.append({'path': relative.as_posix(), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest(), 'originalSource': str(path)})
record = {'sourceRunId': run_id, 'sourceSnapshotId': request['sourceSnapshotId'], 'receiptId': request['receiptId'],
          'sourceRoot': str(source), 'preservedRoot': str(target), 'files': pins, 'fileCount': len(pins), 'totalBytes': sum(item['bytes'] for item in pins), 'mismatches': 0}
with (destination / 'preservation.json').open('x', encoding='utf-8') as stream:
    json.dump(record, stream, indent=2)
    stream.write('\n')
print(json.dumps({key: value for key, value in record.items() if key != 'files'}))
