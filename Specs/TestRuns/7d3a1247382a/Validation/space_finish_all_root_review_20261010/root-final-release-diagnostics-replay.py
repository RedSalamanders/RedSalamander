import csv
import hashlib
import json
from pathlib import Path

repo = Path('C:/Users/PVB/.codex/worktrees/viewer-space-explorer-analyst/RedSalamander')
outroot = Path('D:/RedSalamander.Perf/evidence/viewer-space-finish-all-20261009')
assert (outroot.parents[1] / '.red-salamander-test-root').read_text().strip() == 'red-salamander.test-sandbox-root.v1'
base = repo / 'Specs/TestRuns/7d3a1247382a'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
archives = ['Validation/space_prepush_release_failure_reports_20261009T213050Z', 'Validation/space_prepush_release_viewer_diagnostics_20261009T213050Z', 'Build/space_baseline_release_matched_final_fb45_20261009T223611Z']
results = []
for name in archives:
    arc = base / name
    rows = list(csv.DictReader((arc / 'manifest.tsv').open(encoding='utf-8-sig'), delimiter='\t'))
    available = 0
    for row in rows:
        p = arc / row['path']
        assert p.stat().st_size == int(row['bytes']) and sha(p) == row['sha256']
        source = Path(row.get('preserved_source_path') or row['absolute_source'])
        assert source.stat().st_size == int(row['bytes']) and sha(source) == row['sha256']
        original = row.get('original_source_path', '')
        if original and Path(original).is_file():
            assert sha(Path(original)) == row['sha256']
            available += 1
    results.append({'archive': name, 'pairs_verified': len(rows), 'surviving_original_matches': available})
diagnostics = base / archives[1]
native = diagnostics / 'checkpoint/viewer-pe.noninteractive/results/ViewerPETests.Noninteractive.output.log'
text = native.read_text(encoding='utf-8-sig')
cases = ['TestViewerSpaceScanPolicyHelpers', 'TestViewerSpaceAnalysisHostVirtualGridAndLifetime', 'TestViewerSpaceLifecycleCaptures', 'TestViewerSpaceListingValidation', 'TestViewerSpaceBoundedHostileProviderScanning', 'TestViewerSpaceBlockedProviderCloseAndPostUpdateRace', 'TestViewerSpaceDirectorySummaryProtocol']
for case in cases:
    for witness in ['fresh harness process reports an exit code', 'fresh harness run passes', 'owned job has no active processes before desktop restoration']:
        assert '[PASS] ' + case + ' ' + witness in text
interactive = diagnostics / 'ViewerPETests.Interactive.output.log'
text = interactive.read_text(encoding='utf-8-sig')
case = 'TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses'
witness_counts = {}
for witness in ['fresh harness process reports an exit code', 'fresh harness run passes', 'owned job has no active processes before desktop restoration', 'parent observes exact saved foreground and keyboard focus without restoration writes']:
    witness_counts[witness] = text.count('[PASS] ' + case + ' ' + witness)
    assert witness_counts[witness] == 4
assert sum(line.startswith('[FAIL]') for line in text.splitlines()) == 5
baseline = json.loads((base / archives[2] / 'RedSalamander.exe.DxUi.json').read_text(encoding='utf-8-sig'))
candidate = json.loads((base / 'Build/space_candidate_release_prepush_reattest_20261009T214710Z/RedSalamander.exe.DxUi.json').read_text(encoding='utf-8-sig'))
for key in ['repository', 'commit', 'apiRevision', 'platform', 'configuration', 'publicHeadersGitTree', 'buildIdentity', 'archiveSha256']:
    assert baseline[key] == candidate[key]
record = {'reviewer': 'root', 'archives': results, 'native_noninteractive_cases': cases, 'release_interactive_space_witnesses': witness_counts, 'later_broad_interactive_failures': 5, 'matched_dxui_archive_sha256': baseline['archiveSha256'], 'limits': 'Scoped ViewerSpace witnesses do not replace failed whole PE/PrePush verdicts; repeated churn successes are not independent samples.'}
path = outroot / 'root-final-release-diagnostics-baseline-replay.json'
path.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'record': str(path), 'pairs_verified': sum(r['pairs_verified'] for r in results), 'native_cases': len(cases), 'space_interactive': 'PASSED', 'matched_dxui': True}))
