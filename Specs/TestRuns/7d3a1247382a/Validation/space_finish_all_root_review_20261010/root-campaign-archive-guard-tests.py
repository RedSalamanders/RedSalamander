import dataclasses
import hashlib
import importlib.util
import json
import pathlib
import sys
import uuid

root = pathlib.Path('C:/Users/PVB/.codex/worktrees/viewer-space-explorer-analyst/RedSalamander')
helper = root / '.build/codex-proposals/archive-current-space-campaign.py'
spec = importlib.util.spec_from_file_location('space_archive_root_review', helper)
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
spec.loader.exec_module(module)
module._check_marker_first()
evidence = pathlib.Path('D:/RedSalamander.Perf/evidence/viewer-space-finish-all-20261009')
scratch = evidence / ('root-curator-guard-' + uuid.uuid4().hex)
assert scratch.parent == evidence and scratch.is_relative_to(pathlib.Path('D:/RedSalamander.Perf'))
scratch.mkdir()
checks = []

def rejected(name, action):
    try:
        action()
    except module.ArchiveError as exc:
        checks.append({'name': name, 'status': 'PASS', 'reason': str(exc)})
    else:
        raise AssertionError(name + ' unexpectedly accepted')

try:
    rejected('metadata-inclusive cap', lambda: module._require_archive_size('test', module.MAX_ARCHIVE_PAYLOAD_BYTES, module.MAX_ARCHIVE_METADATA_BYTES + 1))
    existing = scratch / 'existing'
    existing.mkdir()
    rejected('existing destination', lambda: module._require_target_absent(existing, scratch))
    manifest = 'D:/RedSalamander.Perf/evidence/viewer-space-memory-cachedreopen200-20261009T224059615Z-9e53a14b/campaign-manifest.json'
    plan = module.build_plan(manifest, 'cached-200-224059615z-9e53a14b', root)
    leaf = plan.leaves[1]
    collision = dataclasses.replace(leaf.copy_items[0], destination=pathlib.PurePosixPath('README.md'))
    bad_leaf = dataclasses.replace(leaf, copy_items=[collision])
    rejected('copied/generated metadata collision', lambda: module._metadata_payloads(plan, bad_leaf))
    preview = module.preview_plan(plan)
    assert preview['would_write'] is False and len(preview['archives']) == 2
    assert all(item['total_bytes_estimate'] <= module.MAX_ARCHIVE_BYTES for item in preview['archives'])
    assert all(not (plan.target_base / item['folder']).exists() for item in preview['archives'])
    checks.append({'name': 'bounded direct-child preview without writes', 'status': 'PASS'})
finally:
    module._remove_stage_contained(scratch, evidence)

record = {'reviewer': 'root', 'helperSha256': hashlib.sha256(helper.read_bytes()).hexdigest(), 'checks': checks, 'actualCampaignWrites': 0}
(evidence / 'root-campaign-archive-guard-tests.json').write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
print(json.dumps(record))
