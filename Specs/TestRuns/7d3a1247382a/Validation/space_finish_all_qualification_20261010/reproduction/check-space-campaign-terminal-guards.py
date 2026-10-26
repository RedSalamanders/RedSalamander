import copy
import hashlib
import importlib.util
import json
import pathlib
import sys

helper = pathlib.Path(__file__).with_name('archive-current-space-campaign.py')
spec = importlib.util.spec_from_file_location('space_terminal_guard_review', helper)
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
spec.loader.exec_module(module)
module._check_marker_first()
campaign = pathlib.Path('D:/RedSalamander.Perf/evidence/viewer-space-memory-retainedreference-20261009T231508122Z-399289c9')
manifest = json.loads((campaign / 'campaign-manifest.json').read_bytes())
base = copy.deepcopy(manifest)
base['runs'] = [copy.deepcopy(next(run for run in base['runs'] if run['status'] == 'PASSED'))]
base['finished_utc'] = None
results = []

def check(name, value, accepted):
    try:
        _, scope = module._check_campaign_schema(value, campaign)
        result = {'case': name, 'accepted': True, 'scope': scope}
    except module.ArchiveError as error:
        result = {'case': name, 'accepted': False, 'reason': str(error)}
    assert result['accepted'] == accepted, result
    results.append(result)

check('completed-prefix-without-terminal-campaign', base, False)
complete = copy.deepcopy(base)
complete['finished_utc'] = '2026-10-10T00:00:00Z'
check('terminal-passed-campaign', complete, True)
for state in ['STARTING', 'CURATION_PENDING', 'UNKNOWN']:
    altered = copy.deepcopy(complete)
    altered['runs'][0]['status'] = state
    check('reject-primary-' + state, altered, False)
pending = copy.deepcopy(complete)
pending['runs'][0]['case_status'] = 'pending'
check('passed-primary-with-pending-case', pending, False)
failed = copy.deepcopy(base)
failed['runs'][0]['status'] = 'FAILED'
failed['runs'][0]['case_status'] = 'pending'
failed['runs'][0]['finished_utc'] = '2026-10-10T00:00:00Z'
check('sealed-failure-with-unavailable-case', failed, True)
failed['runs'][0]['finished_utc'] = None
check('unsealed-failure-with-pending-case', failed, False)
record = {'reviewer': 'root', 'scope': 'Private curation state guards; not application behavior or a benchmark result',
          'helperSha256': hashlib.sha256(helper.read_bytes()).hexdigest(), 'cases': results, 'failures': 0}
target = pathlib.Path('D:/RedSalamander.Perf/evidence/viewer-space-finish-all-20261009/root-current-campaign-terminal-guard-tests.json')
with target.open('x', encoding='utf-8') as stream:
    json.dump(record, stream, indent=2)
    stream.write('\n')
print(json.dumps({'guardCases': len(results), 'failures': 0, 'record': str(target)}))
