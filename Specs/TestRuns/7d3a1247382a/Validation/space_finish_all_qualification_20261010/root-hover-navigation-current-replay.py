import hashlib
import json
import math
import pathlib
import statistics

authority = pathlib.Path('D:/RedSalamander.Perf')
assert (authority / '.red-salamander-test-root').read_text().strip() == 'red-salamander.test-sandbox-root.v1'
evidence = authority / 'evidence/viewer-space-finish-all-20261009'
preserved = evidence / 'hover-navigation-current'
root = preserved / 'artifacts/selftest/last_run'
report = json.loads((evidence / 'root-hover-navigation-current-replay.json').read_bytes())
preservation = json.loads((preserved / 'preservation.json').read_bytes())
for item in preservation['files']:
    path = (preserved / 'artifacts' / item['path']).resolve(strict=True)
    assert path.is_relative_to((preserved / 'artifacts').resolve())
    data = path.read_bytes()
    assert len(data) == item['bytes'] and hashlib.sha256(data).hexdigest() == item['sha256']
raw_path = root / 'perf/perf_metrics.jsonl'
raw = raw_path.read_bytes()
raw_lines = raw.splitlines(keepends=True)
rows = [json.loads(line) for line in raw_lines]
hover = json.loads((root / 'perf/viewer_space_input_present_metrics.json').read_bytes())['samplesUs']
navigation = json.loads((root / 'perf/viewer_space_navigation_present_metrics.json').read_bytes())
definitions = [('hoverHandlerToNextPresentSubmission', hover),
               ('retainedControlWheelZoomInToTargetPresentSubmission', navigation['zoomInUs']),
               ('retainedControlWheelZoomOutToTargetPresentSubmission', navigation['zoomOutUs'])]
for key, values in definitions:
    assert len(values) == 200
    ordered = sorted(values)
    expected = report['samples'][key]
    assert expected['count'] == 200 and expected['medianUs'] == statistics.median(values)
    assert expected['p95NearestRankUs'] == ordered[189] and expected['maxUs'] == ordered[-1] and expected['minUs'] == ordered[0]
for detail, array in [('wheel-in', navigation['zoomInUs']), ('wheel-out', navigation['zoomOutUs'])]:
    observed = [row for row in rows if row.get('metric') == 'viewer.space.input.fixture_navigation_present_us' and row.get('detail') == detail]
    assert len(observed) == 200 and all(row['hr'] == 0 for row in observed)
    assert [row['durationUs'] for row in observed] == array
metrics = set(report['uiPreparationDiagnostics']) | {'viewer.space.input.fixture_navigation_present_us', 'viewer.space.input.fixture_hover_p95_us'}
for name, summary in report['uiPreparationDiagnostics'].items():
    observed = [row['durationUs'] for row in rows if row.get('metric') == name]
    assert len(observed) == summary['count']
    if observed:
        ordered = sorted(observed)
        assert summary['medianUs'] == statistics.median(observed)
        assert summary['p95NearestRankUs'] == ordered[math.ceil(.95 * len(ordered)) - 1]
        assert summary['maxUs'] == ordered[-1]
numbers = [index + 1 for index, row in enumerate(rows) if row.get('metric') in metrics]
selected = b''.join(raw_lines[index - 1] for index in numbers)
assert 0 < len(selected) <= 2 * 1024 * 1024
retained = evidence / 'root-hover-navigation-current-metrics.jsonl'
with retained.open('xb') as stream:
    stream.write(selected)
mapping = {'source': str(raw_path), 'sourceBytes': len(raw), 'sourceSha256': hashlib.sha256(raw).hexdigest(),
           'sourceRows': len(raw_lines), 'retainedRows': len(numbers), 'retainedBytes': len(selected),
           'retainedSha256': hashlib.sha256(selected).hexdigest(), 'retainedSourceLineNumbers': numbers,
           'selectedMetricNames': sorted(metrics)}
with (evidence / 'root-hover-navigation-current-metric-map.json').open('x', encoding='utf-8') as stream:
    json.dump(mapping, stream, indent=2)
    stream.write('\n')
assert len(retained.read_bytes().splitlines(keepends=True)) == len(numbers)
record = {'reviewer': 'root', 'mismatches': 0, 'independentlyCheckedHoverAndZoomSamples': 600,
          'exactPreservedFiles': len(preservation['files']), 'independentlyCheckedUiMetricFamilies': len(report['uiPreparationDiagnostics']),
          'exactRetainedRows': len(numbers), 'exactRetainedBytes': len(selected), 'sourceBytes': len(raw), 'sourceRows': len(rows),
          'scope': 'Recomputed all600 sample statistics and all reported UI timing families from preserved artifacts; retained each source observation for the reported metric families byte-for-byte. Filtered v1 report plus frozen workspace/receipt, no broader verdict or baseline gain.'}
with (evidence / 'root-hover-navigation-current-independent-replay.json').open('x', encoding='utf-8') as stream:
    json.dump(record, stream, indent=2)
    stream.write('\n')
print(json.dumps(record))
