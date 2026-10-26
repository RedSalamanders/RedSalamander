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
names = [
    'queued-path-frozen-source.json',
    'root-debug-prepush-build-archive-replay.json',
    'root-debug-prepush-three-archive-replay.json',
    'root-debug-checkpoint-diagnostic-archives-replay.json',
    'root-debug-native-child-archive-replay.json',
    'root-debug-current-lifecycle-visual-review.json',
    'root-debug-current-theme-boundary-visual-review.json',
    'root-debug-additional-inflight-visual-review.json',
    'root-release-commands-capture-review.json',
    'root-release-native-reports-build-and-debug-focus-replay.json',
    'root-final-release-diagnostics-baseline-replay.json',
    'root-final-release-diagnostics-replay.py',
    'root-release-capture-review.py',
    'root-release-evidence-replay.py',
    'root-focused-native-launcher-guard-tests.json',
    'root-campaign-archive-guard-tests.py',
    'root-campaign-archive-guard-tests.json',
    'root-dxui-analyzer-guard-review.log',
    'root-dxui-launcher-guard-review.log',
    'root-prepush-curator-review.log',
    'space-doc-capture-preview-20261010-final.log',
]
sources = [(evidence / name, name) for name in names]
sources += [(repo / 'docs/res/viewer-space-captures.provenance.json', 'root-doc-provenance-before-final.json')]
helpers = ['archive-current-space-campaign.py', 'refresh-viewer-space-doc-captures.ps1', 'archive-space-root-review.py', 'curate-viewer-space-prepush-run.py', 'curate-debug-prepush-diagnostics.py', 'curate-release-prepush-diagnostics.py']
sources += [(pathlib.Path(__file__).with_name(name), 'reproduction/' + name) for name in helpers]
base = guard._check_existing_ancestors(repo / 'Specs/TestRuns/7d3a1247382a/Validation')
target = base / 'space_finish_all_root_review_20261010'
guard._require_target_absent(target, base)
items = []
destinations = set()
for source, destination in sources:
    allowed = evidence if source.is_relative_to(evidence) else repo
    guard._safe_source_path(source, allowed)
    guard._add_file_item(items, destinations, source, destination, 'root-review')
readme = b'''# ViewerSpace root review and reproduction evidence\n\nThese records preserve the root agent's independent byte/hash/metric replay and original-resolution visual reviews for the exact named frozen9add builds and runs. They do not promote failed broader Debug/Release PrePush outcomes, establish unmeasured performance targets, or qualify other DPI/locale/platform/provider/human accessibility profiles.\n\nThe pre-refresh documentation provenance is an exact byte copy. Reproduction helpers are private session tooling, archived as reviewed inputs rather than new Tools commands. Their source-run and sandbox paths are historical evidence bindings. Full original metric sources and run-preservation files stay at their exact marked-root paths; unavailable old run paths are not reconstructed.\n'''
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
