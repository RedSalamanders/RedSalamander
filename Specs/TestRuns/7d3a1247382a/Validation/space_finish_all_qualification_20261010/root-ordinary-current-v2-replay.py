import csv
import hashlib
import json
import pathlib
import statistics

repo = pathlib.Path('C:/Users/PVB/.codex/worktrees/viewer-space-explorer-analyst/RedSalamander')
evidence_root = pathlib.Path('D:/RedSalamander.Perf/evidence/viewer-space-finish-all-20261009')
assert pathlib.Path('D:/RedSalamander.Perf/.red-salamander-test-root').read_text().strip() == 'red-salamander.test-sandbox-root.v1'
base = repo / 'Specs/TestRuns/7d3a1247382a/Viewers'
prefix = 'ordinary-current-224729245z-4221bb43-v2'
analysis = json.loads((base / (prefix + '-overview') / 'analysis.json').read_bytes())
measurements = {}
copied = generated = raw_rows = full_bytes = 0
records = []
for archive in sorted(base.glob(prefix + '-*')):
    for entry in csv.DictReader((archive / 'manifest.txt').open(encoding='utf-8-sig', newline=''), delimiter='\t'):
        path = archive / entry['path']
        assert path.resolve().is_relative_to(archive.resolve())
        data = path.read_bytes()
        assert len(data) == int(entry['bytes']) and hashlib.sha256(data).hexdigest() == entry['sha256']
        if entry['source'].startswith('curator:'):
            generated += 1
        else:
            assert data == pathlib.Path(entry['source']).read_bytes()
            copied += 1
    assert sum(p.stat().st_size for p in archive.rglob('*') if p.is_file()) <= 5 * 1024 * 1024
    index = json.loads((archive / 'archive-index.json').read_bytes())
    if index['archive_kind'] != 'run':
        continue
    run = index['runs'][0]
    source_index = json.loads((archive / 'metrics/source-index.json').read_bytes())
    mapping = json.loads((archive / 'raw-metric-replay-map.json').read_bytes())
    assert mapping['schema'] == 'viewer-space.ordinary-current-raw-replay-map.v1'
    assert mapping['source_index']['archive_path'] == 'metrics/source-index.json'
    index_data = (archive / mapping['source_index']['archive_path']).read_bytes()
    assert len(index_data) == mapping['source_index']['bytes']
    assert hashlib.sha256(index_data).hexdigest() == mapping['source_index']['sha256']
    evidence_data = (archive / mapping['evidence']['archive_path']).read_bytes()
    assert mapping['evidence']['archive_path'] == 'evidence.json'
    assert len(evidence_data) == mapping['evidence']['bytes']
    assert hashlib.sha256(evidence_data).hexdigest() == mapping['evidence']['sha256']
    evidence = json.loads(evidence_data)
    assert mapping['source_index']['source_path'] == evidence['rawSourceIndex']['path']
    assert source_index['shards'] == evidence['rawSourceShards']
    assert len(mapping['fragments']) == len(source_index['shards'])
    for ordinal, (fragment, shard) in enumerate(zip(mapping['fragments'], source_index['shards']), 1):
        assert fragment['ordinal'] == ordinal
        assert fragment['source_path'] == shard['path']
        assert fragment['bytes'] == shard['bytes'] and fragment['sha256'] == shard['sha256']
        assert fragment['archive_path'] == 'metrics/' + pathlib.PurePosixPath(shard['path']).name + '.part'
        assert (archive / fragment['archive_path']).resolve().is_relative_to(archive.resolve())
    assert mapping['source_path'] == source_index['source_path']
    assert mapping['source_bytes'] == source_index['source_bytes'] == evidence['sourceBytes']
    assert mapping['source_sha256'] == source_index['source_sha256'] == evidence['sourceSha256']
    assert mapping['source_rows'] == evidence['sourceRows']
    assert mapping['machine_hash'] == evidence['machineHash']
    parts = []
    for shard in source_index['shards']:
        part = (archive / 'metrics' / (pathlib.PurePosixPath(shard['path']).name + '.part')).read_bytes()
        assert len(part) == shard['bytes'] and hashlib.sha256(part).hexdigest() == shard['sha256']
        parts.append(part)
    payload = b''.join(parts)
    proof = json.loads((archive / 'evidence.json').read_bytes())
    assert len(payload) == source_index['source_bytes'] == proof['sourceBytes']
    assert hashlib.sha256(payload).hexdigest() == source_index['source_sha256'] == proof['sourceSha256']
    lines = payload.splitlines()
    assert len(lines) == proof['sourceRows']
    rows = [json.loads(line) for line in lines]
    assert {row['machineHash'] for row in rows if row.get('machineHash')} == {mapping['machine_hash']}
    raw_rows += len(rows)
    full_bytes += len(payload)
    cases = json.loads((archive / 'commands/results.json').read_bytes())['cases']
    assert [case['name'] for case in cases] == ['viewer_space_perf_small_progressive', 'viewer_space_perf_large_optin']
    assert all(case['status'] == 'passed' for case in cases)
    totals = [row for row in rows if row.get('metric') == 'viewer.space.scan.total_us' and row.get('detail') == 'done']
    assert len(totals) == 2
    scans = {int(row['value0']): int(row['durationUs']) for row in totals}
    assert set(scans) == {10000, 100000}
    first_maps = [row for row in rows if row.get('metric') == 'viewer.space.open_to_first_meaningful_paint_us' and row.get('detail') == 'published-map']
    if run['side'] == 'candidate':
        assert len(first_maps) == 2 and all(row['hr'] == 0 for row in first_maps)
        for i, scenario in enumerate(['small_progressive', 'large_optin']):
            artifact = json.loads((archive / 'perf' / ('viewer_space_perf_' + scenario + '_metrics.json')).read_bytes())
            assert first_maps[i]['durationUs'] == artifact['first_meaningful_paint_us']
    record = {'runId': run['run_id'], 'side': run['side'], 'role': run['role'], 'block': run['block'], 'position': run['position'], 'sourceRows': len(rows), 'scanUs': scans, 'firstMapUs': [row['durationUs'] for row in first_maps]}
    records.append(record)
    if run['role'] == 'measurement':
        measurements[(run['block'], run['position'])] = record
assert len(records) == 14 and len(measurements) == 12
ratios = {'small': [], 'large': []}
for block in range(1, 4):
    assert [measurements[(block, p)]['side'] for p in [1,2,3,4]] == ['baseline','candidate','candidate','baseline']
    for a, b in [(1,2),(4,3)]:
        for scenario, count in [('small',10000),('large',100000)]:
            ratios[scenario].append(measurements[(block,b)]['scanUs'][count] / measurements[(block,a)]['scanUs'][count])
for scenario in ratios:
    assert ratios[scenario] == analysis['paired_ratio_summaries'][scenario]['scan_total_us']['ratios']
    assert statistics.median(ratios[scenario]) == analysis['paired_ratio_summaries'][scenario]['scan_total_us']['median']
for i, scenario in enumerate(['small','large']):
    timings = [run['firstMapUs'][i] for run in measurements.values() if run['side'] == 'candidate']
    reported = analysis['candidate_only_first_meaningful_paint']['scenarios'][scenario]['candidate']
    assert statistics.median(timings) == reported['median'] and max(timings) == reported['p95_nearest_rank']
    assert sum(value <= 250000 for value in timings) == reported['target_met_count']
record = {'reviewer': 'root', 'copiedSourceArchivePairs': copied, 'generatedMetadataHashes': generated, 'exactReconstructedRawRows': raw_rows, 'exactReconstructedRawBytes': full_bytes, 'runs': records, 'pairedScanRatios': ratios, 'mismatches': 0, 'scope': '14 runs with two excluded warmups, six adjacent ABBA pairs, exact 10k/100k ordinary fixtures; no stable tail or generalized gain claim'}
with (evidence_root / 'root-ordinary-current-v2-replay.json').open('x', encoding='utf-8') as output:
    output.write(json.dumps(record, indent=2) + '\n')
print(json.dumps({key: value for key, value in record.items() if key not in {'runs','pairedScanRatios'}}))
