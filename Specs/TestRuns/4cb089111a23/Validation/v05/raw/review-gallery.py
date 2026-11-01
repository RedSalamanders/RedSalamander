import hashlib
import json
from pathlib import Path
from PIL import Image, ImageChops

root = Path(__file__).parent
previous = root.parent / 'semantic-detail-qualified-v6-20260923'
tail = 'gallery-runner-artifacts/artifacts/selftest/last_run/fileops/visual-gallery'
before = {p.name: p for p in (previous / tail).glob('*.png')}
after = {p.name: p for p in (root / tail).glob('*.png')}
assert len(before) == len(after) == 538
assert before.keys() == after.keys()
results = []
for name in sorted(after):
    a, b = before[name], after[name]
    ah = hashlib.sha256(a.read_bytes()).hexdigest()
    bh = hashlib.sha256(b.read_bytes()).hexdigest()
    with Image.open(a) as ai, Image.open(b) as bi:
        dimensions_equal = ai.size == bi.size
        pixels_equal = dimensions_equal and not any(high for low, high in ImageChops.difference(ai.convert('RGBA'), bi.convert('RGBA')).getextrema())
        results.append(dict(name=name, before_sha256=ah, after_sha256=bh,
                            before_size=ai.size, after_size=bi.size,
                            dimensions_equal=dimensions_equal, pixels_equal=pixels_equal))
(root / 'capture-comparison.json').write_text(json.dumps(results, indent=2) + '\n', encoding='utf-8')
changed = [r for r in results if not r['pixels_equal']]
print(json.dumps(dict(paired=len(results), pixel_changes=len(changed),
                     dimension_changes=sum(not r['dimensions_equal'] for r in results),
                     changed_names=[r['name'] for r in changed]), indent=2))
for case in ('batch', 'recycle', 'gallery'):
    entries = json.loads((root / (case + '-runner-copy-hashes.json')).read_text(encoding='utf-8-sig'))
    for entry in entries:
        p = root / (case + '-runner-artifacts') / entry['path'].replace('\\', '/')
        assert hashlib.sha256(p.read_bytes()).hexdigest() == entry['sha256'], p
    print(f'{case}: {len(entries)} copied runner files verified')
