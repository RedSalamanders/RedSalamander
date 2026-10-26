import hashlib
import json
import pathlib

repo = pathlib.Path(__file__).resolve().parents[2]
authority = pathlib.Path('D:/RedSalamander.Perf')
assert (authority / '.red-salamander-test-root').read_text().strip() == 'red-salamander.test-sandbox-root.v1'
root = authority / 'evidence/viewer-space-finish-all-20261009'
frozen = json.loads((root / 'queued-path-frozen-source.json').read_bytes())
assert frozen['SnapshotId'] == '9addaf53d4d88d618223a7f2d8554ac0b73f2472946c3a3e0875b1fd94ac9554'
paths = ['Common/WindowMessages.h', 'Plugins/ViewerSpace/ViewerSpace.Layout.cpp',
         'Plugins/ViewerSpace/ViewerSpace.cpp', 'Plugins/ViewerSpace/ViewerSpace.h',
         'RedSalamander/SelfTest/Commands/Commands.Tests.ViewCommands.cpp']
code = []
for path in paths:
    data = (repo / path).read_bytes()
    pin = next(item for item in frozen['WorkspaceSnapshot']['files'] if item['path'] == path)
    assert len(data) == pin['worktree_size_bytes']
    assert hashlib.sha256(data).hexdigest() == pin['worktree_sha256']
    code.append({'path': path, 'bytes': len(data), 'sha256': pin['worktree_sha256']})

def pinned_json(path):
    data = path.read_bytes()
    return json.loads(data), {'path': str(path), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}

reports = {}
for profile, timestamp, receipt in [
    ('Debug', '201507', 'f721d4c3644b36016e3b81900978cc335f1a530d9eadc094f585683bc7f3638b'),
    ('Release', '213050', '37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd')]:
    path = repo / f'Specs/TestRuns/7d3a1247382a/Validation/space_{profile.lower()}_prepush_reports_20261009T{timestamp}Z/selftest/last_run/run-all-tests-results.json'
    report, pin = pinned_json(path)
    assert report['schema'] == 'red-salamander.run-all-tests.v2'
    assert report['build_receipt_id'] == receipt
    assert report['snapshot_id'] == '430becd60c8af4bdcb56d4034e78e20028381748b4804249530bb9d0f3ee301e'
    workspace, workspace_pin = pinned_json(path.with_name('workspace-final.json'))
    assert workspace['snapshot_id'] == workspace['acquisition_before'] == workspace['acquisition_after'] == report['snapshot_id']
    for frozen_file in frozen['WorkspaceSnapshot']['files']:
        observed = next(item for item in workspace['files'] if item['path'] == frozen_file['path'])
        assert observed['worktree_sha256'] == frozen_file['worktree_sha256']
        assert observed['worktree_size_bytes'] == frozen_file['worktree_size_bytes']
    assert report['repository_verdict'] == report['change_verdict'] == 'FAILED'
    reports[profile] = {'source': pin, 'runId': report['run_id'], 'receiptId': receipt,
                        'testSnapshotId': report['snapshot_id'], 'workspaceFinal': workspace_pin,
                        'verifiedFrozenWorkspaceFiles': len(frozen['WorkspaceSnapshot']['files']),
                        'caseCounts': report['case_counts'], 'entryCounts': report['counts'],
                        'repositoryVerdict': report['repository_verdict'], 'changeVerdict': report['change_verdict'],
                        'entries': [{'id': item['id'], 'promotionState': item['promotion_state'],
                                     'execution': item['execution'], 'reasonCodes': item['reason_codes']} for item in report['entries']]}

campaign_names = {
    'ordinary': 'viewer-space-ordinary-current-20261009T224729245Z-4221bb43',
    'cached200': 'viewer-space-memory-cachedreopen200-20261009T224059615Z-9e53a14b',
    'candidateHeavy': 'viewer-space-memory-candidateheavy-20261009T224302601Z-0b9ae69c',
    'controlledRetained': 'viewer-space-memory-retainedreference-20261009T231508122Z-399289c9',
}
campaigns = {}
for key, name in campaign_names.items():
    folder = authority / 'evidence' / name
    manifest, manifest_pin = pinned_json(folder / 'campaign-manifest.json')
    assert manifest['finished_utc']
    assert all(run['status'] == 'PASSED' for run in manifest['runs'])
    analysis, analysis_pin = pinned_json(folder / 'analysis.json')
    campaigns[key] = {'manifest': manifest_pin, 'analysis': analysis_pin,
                      'campaignId': manifest['campaign_id'], 'runCount': len(manifest['runs']),
                      'runIds': [run['run_id'] for run in manifest['runs']], 'result': analysis}

input_latency, input_pin = pinned_json(root / 'root-hover-navigation-current-replay.json')
independent, independent_pin = pinned_json(root / 'root-hover-navigation-current-independent-replay.json')
assert input_latency['mismatches'] == independent['mismatches'] == 0
assert input_latency['provenance']['sourceSnapshotId'] == frozen['SnapshotId']
assert input_latency['provenance']['receiptId'] == '37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd'
assert input_latency['provenance']['caseStatus'] == 'passed'
assert independent['independentlyCheckedHoverAndZoomSamples'] == 600

matrix = {'schema': 'viewer-space.root-final-qualification.v1', 'reviewer': 'root',
          'qualifiedBuildSource': frozen['SnapshotId'], 'verifiedCodeFiles': code,
          'qualificationScope': 'Named frozen-source x64 Debug/Release runs only; subsequent documentation/evidence changes do not make these receipts qualification of a new complete workspace snapshot.',
          'prePush': reports, 'campaigns': campaigns,
          'currentInputLatency': {'source': input_pin, 'independentReplay': independent_pin, 'result': input_latency},
          'absoluteTargetsUnchanged': {'ordinaryFirstMapUs': 250000, 'cachedFirstMapUs': 100000,
                                       'historicalS0PeakPrivateDeltaRatio': 0.70},
          'historicalS0Status': 'UNMEASURED: controlled same-IViewer reference has a different explicitly labeled lifecycle denominator.',
          'unresolved': [
              {'gate': 'cached-first-map', 'owner': 'user and ViewerSpace maintainer', 'status': 'MISS; section 12 disposition pending; no waiver'},
              {'gate': 'small-ordinary-first-map', 'owner': 'ViewerSpace maintainer', 'status': 'MISS; 0 of 6 under 250 ms'},
              {'gate': 'broad-PrePush', 'owner': 'respective Commands, Compare, FileOps, provider and native test owners', 'status': 'FAILED; preserved original results, no inferred repairs'},
              {'gate': 'Windows symbolic links', 'owner': 'Windows test-environment owner', 'status': 'privilege unavailable; no environment changes authorized'},
              {'gate': 'hosted CI', 'owner': 'repository/account billing owner', 'status': 'published prior head failed before steps due to billing; refresh new head once after push'},
              {'gate': 'native ARM64', 'owner': 'native device/runner owner', 'status': 'not qualified'},
              {'gate': 'live providers/devices', 'owner': 'provider account/device owner', 'status': 'not qualified'},
              {'gate': 'translation and broader DPI', 'owner': 'translation/UI reviewer', 'status': 'current actual captures qualify named 144-DPI scenarios only'},
              {'gate': 'human screen-reader review', 'owner': 'accessibility reviewer', 'status': 'pending'},
              {'gate': 'wide/skewed/provider tails and general no-regression', 'owner': 'performance maintainer', 'status': 'not qualified by ordinary or controlled cohorts'},
          ], 'planStatus': 'ACTIVE', 'targetWaivers': []}
destination = root / 'root-final-current-qualification-matrix.json'
with destination.open('x', encoding='utf-8') as stream:
    json.dump(matrix, stream, indent=2)
    stream.write('\n')
print(json.dumps({'record': str(destination), 'verifiedCodeFiles': len(code),
                  'campaigns': {key: value['runCount'] for key, value in campaigns.items()},
                  'repositoryVerdicts': {key: value['repositoryVerdict'] for key, value in reports.items()}}))
