import csv
import hashlib
import json
import pathlib
import statistics

repo = pathlib.Path(__file__).resolve().parents[2]
authority = pathlib.Path('D:/RedSalamander.Perf')
assert (authority / '.red-salamander-test-root').read_text().strip() == 'red-salamander.test-sandbox-root.v1'
campaign_root = authority / 'evidence/viewer-space-memory-retainedreference-20261009T231508122Z-399289c9'
evidence = authority / 'evidence/viewer-space-finish-all-20261009'
manifest = json.loads((campaign_root / 'campaign-manifest.json').read_bytes())
analysis = json.loads((campaign_root / 'analysis.json').read_bytes())
assert manifest['finished_utc'] and len(manifest['runs']) == 14
assert all(item['status'] == 'PASSED' for item in manifest['runs'])
assert analysis['status'] == 'qualified'
prefix = 'retained-reference-231508122z-399289c9'
base = repo / 'Specs/TestRuns/7d3a1247382a/Viewers'
expected_leaves = {}
for run in manifest['runs']:
    name = (prefix + '-warmup-' + run['side']) if run['role'] == 'warmup' else f"{prefix}-b{run['block']:02d}-p{run['position']:02d}-{run['side']}"
    assert name not in expected_leaves
    expected_leaves[name] = run
archives = sorted(base.glob(prefix + '-*'))
assert {archive.name for archive in archives} == {*expected_leaves, prefix + '-overview'}
assert len(archives) == 15 and all(item.is_dir() for item in archives)
overview = base / (prefix + '-overview')
assert (overview / 'campaign-manifest.json').read_bytes() == (campaign_root / 'campaign-manifest.json').read_bytes()
assert (overview / 'analysis.json').read_bytes() == (campaign_root / 'analysis.json').read_bytes()
copied = generated = replayed = source_rows = source_bytes = 0
map_count = expected_selected_rows = expected_source_rows = expected_source_bytes = 0
for archive in archives:
    archive_index = json.loads((archive / 'archive-index.json').read_bytes())
    assert archive_index['campaign_id'] == manifest['campaign_id']
    assert archive_index['archive_folder'] == archive.name
    assert archive_index['campaign_manifest_sha256'] == hashlib.sha256((campaign_root / 'campaign-manifest.json').read_bytes()).hexdigest()
    expected_run = expected_leaves.get(archive.name)
    if expected_run:
        assert len(archive_index['runs']) == 1
        archived_run = archive_index['runs'][0]
        assert all(archived_run[field] == expected_run[field] for field in ['run_id', 'side', 'role', 'block', 'position', 'status'])
        assert pathlib.Path(archived_run['source_path']).resolve() == pathlib.Path(expected_run['archive_path']).resolve()
        assert len(list(archive.rglob('metric-replay-map.json'))) == 1
        for name in ['metric-replay-map.json', 'raw-metrics-source-index.json', 'perf-metrics-curated.jsonl', 'provenance.json', 'viewer_space_memory_heavy_four_root_retained_reference_metrics.json']:
            assert (archive / name).read_bytes() == (pathlib.Path(expected_run['archive_path']) / name).read_bytes()
        provenance = json.loads((archive / 'provenance.json').read_bytes())
        assert provenance['runId'] == expected_run['run_id']
        expected_selected_rows += provenance['metrics']['retainedRows']
        expected_source_rows += provenance['rawMetricRowCount']
        expected_source_bytes += provenance['rawMetricsBytes']
    else:
        assert len(archive_index['runs']) == 14 and not list(archive.rglob('metric-replay-map.json'))
    with (archive / 'manifest.txt').open(encoding='utf-8-sig', newline='') as stream:
        entries = list(csv.DictReader(stream, delimiter='\t'))
    names = set()
    for entry in entries:
        relative = pathlib.PurePosixPath(entry['path'])
        assert not relative.is_absolute() and '..' not in relative.parts
        assert str(relative).casefold() not in names
        names.add(str(relative).casefold())
        destination = (archive / str(relative)).resolve(strict=True)
        assert destination.is_relative_to(archive.resolve())
        data = destination.read_bytes()
        assert len(data) == int(entry['bytes']) and hashlib.sha256(data).hexdigest() == entry['sha256']
        if entry['source'].startswith('curator:'):
            generated += 1
        else:
            source = pathlib.Path(entry['source']).resolve(strict=True)
            assert source.is_relative_to(authority.resolve()) or source.is_relative_to(repo.resolve())
            assert data == source.read_bytes()
            copied += 1
    for replay in archive.rglob('metric-replay-map.json'):
        map_count += 1
        mapping = json.loads(replay.read_bytes())
        index = json.loads((replay.parent / 'raw-metrics-source-index.json').read_bytes())
        source = pathlib.Path(index['preserved_path']).resolve(strict=True)
        assert source.is_relative_to(campaign_root.resolve())
        raw = source.read_bytes()
        retained = (replay.parent / 'perf-metrics-curated.jsonl').read_bytes()
        assert hashlib.sha256(raw).hexdigest() == index['sha256'] == mapping['sourceSha256']
        assert len(raw) == index['bytes'] == mapping['sourceBytes']
        raw_lines = raw.splitlines(keepends=True)
        lines = retained.splitlines(keepends=True)
        assert len(raw_lines) == mapping['sourceRows']
        assert hashlib.sha256(retained).hexdigest() == mapping['retainedSha256']
        assert len(lines) == mapping['retainedRows'] == len(mapping['retainedSourceLineNumbers'])
        for line, number in zip(lines, mapping['retainedSourceLineNumbers']):
            assert line == raw_lines[number - 1]
            assert json.loads(line)['machineHash'] == '7d3a1247382a'
        replayed += len(lines)
        source_rows += len(raw_lines)
        source_bytes += len(raw)
assert map_count == len(expected_leaves) == 14
assert replayed == expected_selected_rows > 0
assert source_rows == expected_source_rows > replayed
assert source_bytes == expected_source_bytes > 0

rows = {}
digest = '55EB54080C36C593FCB0AD2454A67F322D4B848995A986A19DDBC3F78FC5D13B'
phases = {'open-loaded', 'seed-closed-object-retained', 'reopened-cached', 'refresh-inflight', 'refresh-done', 'refresh-closed-object-retained'}
phase_count = 0
for run in manifest['runs']:
    name = (prefix + '-warmup-' + run['side']) if run['role'] == 'warmup' else f"{prefix}-b{run['block']:02d}-p{run['position']:02d}-{run['side']}"
    leaf = base / name
    artifact = json.loads((leaf / 'viewer_space_memory_heavy_four_root_retained_reference_metrics.json').read_bytes())
    provenance = json.loads((leaf / 'provenance.json').read_bytes())
    receipt = json.loads((leaf / 'build-receipt.json').read_bytes())
    assert receipt['tests_enabled'] and receipt['configuration'] == 'Release' and receipt['platform'] == 'x64'
    assert provenance['buildReceiptId'] == run['receipt_id'] == receipt['receipt_id']
    assert provenance['workspaceSnapshotId'] == provenance['buildSourceSnapshotId'] == receipt['source_snapshot_id'] == run['workspace_snapshot_id']
    assert artifact['rootCount'] == artifact['cacheHitCount'] == 4
    assert artifact['filesPerRoot'] == 100000 and artifact['directoriesPerRoot'] == 501
    assert artifact['rootTotalBytes'] == 3001292 and artifact['rootFixtureDigests'] == [digest] * 4
    assert artifact['configuration'] == 'topK512_threads8_volume1_cacheTTL600_entries4' and artifact['effectiveScanThreads'] == 8
    assert len(artifact['rootSetupPerRootUs']) == 4 and all(0 <= value <= 120000000 for value in artifact['rootSetupPerRootUs'])
    assert 0 <= artifact['lifecycleUs'] <= 360000000
    assert len(artifact['phases']) == 25
    for root in range(1, 5):
        observed = [phase for phase in artifact['phases'] if phase['root'] == root]
        assert len(observed) == 6 and {phase['phase'] for phase in observed} == phases
        assert all(phase['rootBytes'] == 3001292 for phase in observed if phase['phase'] != 'refresh-inflight')
        inflight = next(phase for phase in observed if phase['phase'] == 'refresh-inflight')
        assert isinstance(inflight['scanState'], int) and 0 <= inflight['scanState'] <= 5
    assert len([phase for phase in artifact['phases'] if phase['root'] == 0 and phase['phase'] == 'final-owner-release']) == 1
    phase_count += 25
    for field in ['privateBytes', 'workingSetBytes']:
        assert artifact['processPeak'][field] >= max(artifact['processBefore'][field], artifact['processAfter'][field])
    if run['role'] == 'measurement':
        rows[(run['block'], run['position'])] = (run, artifact)
assert len(rows) == 12
ratios = []
for block in range(1, 4):
    for pair_index, (baseline_position, candidate_position) in enumerate([(1, 2), (4, 3)], 1):
        baseline, b = rows[(block, baseline_position)]
        candidate, c = rows[(block, candidate_position)]
        assert baseline['side'] == 'baseline' and candidate['side'] == 'candidate'
        b_delta = b['processPeak']['privateBytes'] - b['processBefore']['privateBytes']
        c_delta = c['processPeak']['privateBytes'] - c['processBefore']['privateBytes']
        assert b_delta > 0 and c_delta >= 0
        ratio = c_delta / b_delta
        result = next(item for item in analysis['pairedContrasts'] if item['block'] == block and item['pair'] == pair_index)
        assert result['baselineRunId'] == baseline['run_id'] and result['candidateRunId'] == candidate['run_id']
        assert result['baselinePeakPrivateDeltaBytes'] == b_delta and result['candidatePeakPrivateDeltaBytes'] == c_delta
        assert result['pairedPeakPrivateDeltaRatioCandidateOverBaseline'] == ratio
        ratios.append(ratio)
median = statistics.median(ratios)
assert median == analysis['candidateVsBaselineRatio']['pairedPeakPrivateDeltaRatioMedian']
assert analysis['memoryReductionTarget']['met'] == (median <= .70)
assert 'not historical' in analysis['memoryReductionTarget']['denominator']
record = {'reviewer': 'root', 'mismatches': 0, 'archives': len(archives), 'exactCopiedPayloads': copied,
          'generatedPayloadDigests': generated, 'byteExactReplayedSelectedRows': replayed,
          'requiredReplayMaps': map_count,
          'preservedSourceRows': source_rows, 'preservedSourceBytes': source_bytes,
          'fixtureArtifactWitnesses': 14, 'phaseArtifactWitnesses': phase_count, 'pairedContrasts': 6,
          'pairedPeakPrivateDeltaRatios': ratios, 'pairedPeakPrivateDeltaRatioMedian': median,
          'scope': 'Controlled same-IViewer lifecycle, not historical S0. All selected raw rows replayed byte-for-byte against full preserved source hashes. Equal fixture identity and exact retained detail are harness artifact witnesses, not a separate re-enumeration of fixture files. Refresh first-map values remain Open-scoped.'}
destination = evidence / 'root-retained-reference-current-replay.json'
with destination.open('x', encoding='utf-8') as stream:
    json.dump(record, stream, indent=2)
    stream.write('\n')
print(json.dumps(record))
