import hashlib
import json
from pathlib import Path

root = Path('C:/RedSalamander.Perf/evidence/i26-ui')
out = root / 'common-p0-current-crossover-20260927'
out.mkdir(exist_ok=False)
policies = {
    'A': ('C:/RsI26P0', 'common-p0-baseline-v2-20260927'),
    'B': ('C:/RsI26Final', 'exact-bound-results-candidate-v5-20260927'),
}
sequence = 'ABBAABBAAB'
plan = []
for index, label in enumerate(sequence, 1):
    repo, previous = policies[label]
    d = out / f'{index:02}-{label}'
    d.mkdir()
    source = root / previous
    inputs = json.loads((source / 'candidate-inputs.json').read_text(encoding='utf-8-sig'))
    for item in inputs:
        assert hashlib.sha256((Path(repo) / item['path']).read_bytes()).hexdigest() == item['sha256']
    (d / 'candidate-inputs.json').write_bytes((source / 'candidate-inputs.json').read_bytes())
    driver = (source / 'run-focused.ps1').read_text(encoding='utf-8-sig')
    driver = driver.replace(str(source).replace('\\', '/'), str(d).replace('\\', '/'))
    start = driver.index('$cases = [ordered]@{')
    end = driver.index('\n$receipts =', start)
    driver = driver[:start] + "$cases = [ordered]@{ common_smoke = 'FileOps_CommonPresentationPerf' }" + driver[end:]
    (d / 'run.ps1').write_bytes(driver.encode('utf-8'))
    exe = Path(repo) / '.build/x64/Release/RedSalamander.exe'
    plan.append(dict(ordinal=index,policy=label,repo=repo,directory=str(d),exe_sha256=hashlib.sha256(exe.read_bytes()).hexdigest()))
(out / 'plan.json').write_bytes((json.dumps(plan, indent=2)+'\n').encode('utf-8'))
(out / 'README.md').write_bytes(b'''# Predeclared common-presentation comparison

Five independent processes per revision in ABBAABBAAB order. A is original P0 e61e349c with pin f5c7fc74; B is frozen current V5 with pin 78b3de38. Identical fixture bytes and semantic workload, DPI and dimensions are required. Each invocation uses the governed runner with SkipBuild and full-solution attestation; no foreground input. This compares the whole redesign against P0, not an isolated library optimization. All attempts remain archived, including interrupted or noisy runs. Active compiler/linker/test processes stop dispatch; before/after process inventories document activity. Each complete runner directory is copied and rehashed before the next run.

Report synchronous redraw/churn and stage-relative resources, including CPU and hidden/destruction behavior. No presented-FPS, frame-pacing, surface-allocation, selected-detail parity or final combined-pin acceptance claim. Do not invent performance budgets from these first samples or apply the grid-only memory approval to the consumer. Preserve original timing and same-policy scatter; any confirmed new degradation remains open for analysis/advice.
''')
print(out)
