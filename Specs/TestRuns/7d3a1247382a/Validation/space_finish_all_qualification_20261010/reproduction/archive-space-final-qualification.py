import argparse
import csv
import hashlib
import importlib.util
import io
import json
import os
import pathlib
import sys
import uuid

spec = importlib.util.spec_from_file_location('space_root_archive_guard', pathlib.Path(__file__).with_name('archive-current-space-campaign.py'))
guard = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = guard
spec.loader.exec_module(guard)
guard._check_marker_first()
repo = guard._repo_root_after_marker()
evidence = pathlib.Path('D:/RedSalamander.Perf/evidence/viewer-space-finish-all-20261009')
terminal = pathlib.Path('D:/RedSalamander.Perf/evidence/viewer-space-memory-retainedreference-20261009T231508122Z-399289c9/campaign-manifest.json')
campaign = json.loads(terminal.read_bytes())
assert campaign['finished_utc'] and all(item['status'] == 'PASSED' for item in campaign['runs'])
names = [
    'queued-path-frozen-source.json',
    'root-final-current-qualification-matrix.json',
    'root-ordinary-current-v2-replay.json',
    'root-ordinary-current-v2-replay.py',
    'root-ordinary-v2-archive-validation.log',
    'root-retained-reference-current-replay.json',
    'root-retained-reference-current-replay.py',
    'root-retained-reference-archive-validation.log',
    'root-current-campaign-terminal-guard-tests.json',
    'root-hover-navigation-current-replay.json',
    'root-hover-navigation-current-replay.py',
    'root-hover-navigation-current-independent-replay.json',
    'root-hover-navigation-current-metric-map.json',
    'root-hover-navigation-current-metrics.jsonl',
    'root-final-documentation-structural-review.json',
    'root-rejected-ordinary-v1-preservation.json',
    'root-partial-byte-shards-validation-rejected.log',
    'root-final-documentation-reconciliation.json',
    'root-final-documentation-applied.diff',
    'root-current-documentation-capture-refresh.log',
    'root-current-documentation-capture-replay.json',
    'root-final-spec-inventory.log',
    'root-final-spec-inventory-replay.json',
    'root-final-change-impact-explain.log',
]
sources = [(evidence / name, name) for name in names]
focus = evidence / 'hover-navigation-current'
focus_names = ['request.json', 'preservation.json', 'build-receipt.json', 'runner-stdout.log', 'runner-stderr.log',
               'artifacts/selftest/last_run/results.json',
               'artifacts/selftest/last_run/run-all-tests-results.json',
               'artifacts/selftest/last_run/run-all-tests-case-history.jsonl',
               'artifacts/selftest/last_run/workspace-final.json',
               'artifacts/selftest/last_run/commands/results.json',
               'artifacts/selftest/last_run/commands/trace.txt',
               'artifacts/selftest/last_run/perf/viewer_space_input_present_metrics.json',
               'artifacts/selftest/last_run/perf/viewer_space_navigation_present_metrics.json']
sources += [(focus / name, 'current-input/' + name) for name in focus_names]
sources += [(evidence / 'final-spec-inventory' / name, 'spec-inventory/' + name)
            for name in ['normative-consistency.json', 'spec-inventory.md', 'test-inventory.json']]
sources += [(repo / 'docs/res/viewer-space-captures.provenance.json', 'root-doc-provenance-after-final.json')]
helpers = [
    'archive-current-space-campaign.py',
    'check-space-campaign-terminal-guards.py',
    'refresh-viewer-space-doc-captures.ps1',
    'build-space-final-qualification-matrix.py',
    'reconcile-space-final-docs.py',
    'preserve-space-rejected-ordinary-curation.ps1',
    'archive-space-final-qualification.py',
    'run-space-current-hover-navigation.ps1',
    'preserve-space-current-hover-navigation.py',
    'analyze-space-current-hover-navigation.py',
    'root-current-documentation-capture-replay.py',
    'space-graphics-device-reuse-scope-20261009.md',
    'space-current-prepush-s3-readonly-triage-20261009.md',
    'space-current-prepush-focus-restoration-readonly-triage-20261009.md',
    'space-current-prepush-compare-readonly-triage-20261009.md',
    'space-current-prepush-clipboard-readonly-triage-20261009.md',
]
sources += [(pathlib.Path(__file__).with_name(name), 'reproduction/' + name) for name in helpers]
base = guard._check_existing_ancestors(repo / 'Specs/TestRuns/7d3a1247382a/Validation')
target = base / 'space_finish_all_qualification_20261010'
guard._require_target_absent(target, base)
items = []
destinations = set()
for source, destination in sources:
    allowed = evidence if source.is_relative_to(evidence) else repo
    guard._safe_source_path(source, allowed)
    guard._add_file_item(items, destinations, source, destination, 'root-review')
readme = b'''# ViewerSpace final frozen-source qualification evidence\n\nRoot verified the exact named frozen9add x64 Debug/Release builds, four completed current-source campaigns, all corrected ordinary raw journals, controlled retained-reference fixture/phases and paired memory ratios, and the three harness-produced documentation images. The test receipts qualify their frozen source; later documentation/evidence changes are not qualification of a new complete workspace snapshot.\n\nBoth broader PrePush attempts remain FAILED. Cached100ms and small ordinary250ms remain MISS. Historical fresh-instance memory70pct-of-S0 remains UNMEASURED; the controlled same-IViewer denominator is separately labeled. The plan remains ACTIVE and no target is waived. Current captures qualify named144DPI scenarios only. Human accessibility, broader platform/provider/locale tails, CI billing and Windows symlink-privilege gates retain their separate owners.\n\nReproduction helpers are private session tooling archived as reviewed inputs, not new Tools commands. Source-run and sandbox paths are historical bindings. Rejected initial ordinary byte-fragment curation is preserved outside Specs; corrected v2 retains identical original metrics. Full original journals remain at their exact marked-root paths.\n'''
rows = [{'path': item.destination.as_posix(), 'bytes': item.size, 'sha256': item.sha256, 'source': str(item.source)} for item in items]
for name, payload in [('README.md', readme), ('.gitattributes', b'* -text whitespace=cr-at-eol\n')]:
    rows.append({'path': name, 'bytes': len(payload), 'sha256': hashlib.sha256(payload).hexdigest(), 'source': 'curator:root-review-metadata'})
stream = io.StringIO(newline='')
writer = csv.DictWriter(stream, fieldnames=['path', 'bytes', 'sha256', 'source'], delimiter='\t', lineterminator='\n')
writer.writeheader()
writer.writerows(rows)
manifest = stream.getvalue().encode('utf-8')
metadata = len(readme) + len(b'* -text whitespace=cr-at-eol\n') + 2 * len(manifest)
guard._require_archive_size(target.name, sum(item.size for item in items), metadata)
parser = argparse.ArgumentParser()
parser.add_argument('--execute', action='store_true')
args = parser.parse_args()
if not args.execute:
    print(json.dumps({'status': 'PREVIEW_ONLY', 'target': str(target), 'files': len(items), 'payloadBytes': sum(item.size for item in items), 'metadataBytes': metadata, 'sourcePairs': rows}, indent=2))
    raise SystemExit(0)
stage = base / ('.' + target.name + '.staging-' + uuid.uuid4().hex)
stage.mkdir()
try:
    for item in items:
        guard._copy_exact(item, stage)
    for name, payload in [('README.md', readme), ('.gitattributes', b'* -text whitespace=cr-at-eol\n'), ('manifest.tsv', manifest), ('manifest.txt', manifest)]:
        with (stage / name).open('xb') as output:
            output.write(payload)
    total, count = guard._measure_tree_bytes(stage)
    guard._require_archive_size(target.name, total - metadata, metadata)
    for item in items:
        data = (stage / item.destination).read_bytes()
        assert len(data) == item.size and hashlib.sha256(data).hexdigest() == item.sha256
        assert data == item.source.read_bytes()
    guard._require_target_absent(target, base)
    os.rename(stage, target)
    print(json.dumps({'status': 'ARCHIVED', 'target': str(target), 'files': count, 'bytes': total}))
finally:
    if stage.exists():
        guard._remove_stage_contained(stage, base)
